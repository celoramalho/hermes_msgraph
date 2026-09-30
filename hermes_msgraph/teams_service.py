import time

import requests

from exceptions import HermesMSGraphError


class TeamsService:
    """
    Service to read Microsoft Teams data via Microsoft Graph:
    teams, channels, channel messages (+ replies), user chats,
    chat messages and chat members.

    Requires application permissions (admin consented):
    Team.ReadBasic.All, Channel.ReadBasic.All, ChannelMessage.Read.All,
    Chat.Read.All, ChatMember.Read.All, User.Read.All.
    """

    BASE_URL = "https://graph.microsoft.com/v1.0"
    BETA_URL = "https://graph.microsoft.com/beta"

    def __init__(self, http_client):
        self.http = http_client
        self.HermesMSGraphError = HermesMSGraphError

    def __get_all_pages(self, url, max_retries=5, progress_callback=None):
        """
        Follows @odata.nextLink until every page has been collected,
        retrying on 429 (throttling) using the Retry-After header, and on
        timeouts/connection errors with a short exponential backoff.

        Args:
            progress_callback: optional callable invoked after each page is
                fetched as progress_callback(pages_fetched, items_so_far).
                Useful to report progress on endpoints that paginate over
                many pages before the full list is returned (e.g. a channel
                with a large message history).
        """
        results = []
        next_url = url
        retries = 0
        network_retries = 0
        max_network_retries = 5
        pages_fetched = 0

        while next_url:
            try:
                response = self.http.get(next_url)
            except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as erro:
                if network_retries >= max_network_retries:
                    raise self.HermesMSGraphError(
                        f"Too many network retries fetching {next_url}: {erro}"
                    )
                wait_seconds = 2 ** network_retries
                print(
                    f"[hermes_msgraph] Network error on {next_url} ({erro}). "
                    f"Retrying in {wait_seconds}s (attempt {network_retries + 1}/{max_network_retries})..."
                )
                time.sleep(wait_seconds)
                network_retries += 1
                continue

            network_retries = 0

            if response.status_code == 429:
                if retries >= max_retries:
                    raise self.HermesMSGraphError(
                        f"Too many retries after 429 throttling for {next_url}"
                    )
                retry_after = int(response.headers.get("Retry-After", "5"))
                print(
                    f"[hermes_msgraph] Throttled (429) on {next_url}. "
                    f"Retrying in {retry_after}s (attempt {retries + 1}/{max_retries})..."
                )
                time.sleep(retry_after)
                retries += 1
                continue

            if response.status_code != 200:
                raise self.HermesMSGraphError(
                    f"Error fetching data from {next_url}: {response.status_code} - {response.text}"
                )

            retries = 0
            data = response.json()
            results.extend(data.get("value", []))
            next_url = data.get("@odata.nextLink")
            pages_fetched += 1

            if progress_callback is not None:
                progress_callback(pages_fetched, len(results))

        return results

    # ---------------- Teams / Channels ----------------

    def list_all_teams(self):
        """
        List every team (Microsoft 365 group with a Team provisioned) in the tenant.
        """
        url = f"{self.BASE_URL}/teams"
        return self.__get_all_pages(url)

    def list_joined_teams_by_user_id(self, user_id):
        """
        List teams a specific user is a member of.
        """
        url = f"{self.BASE_URL}/users/{user_id}/joinedTeams"
        return self.__get_all_pages(url)

    def list_team_members(self, team_id):
        """
        List members of a team, via the underlying Microsoft 365 group.
        Requires GroupMember.Read.All or Group.Read.All (application permission).
        """
        url = f"{self.BASE_URL}/groups/{team_id}/members"
        return self.__get_all_pages(url)

    def list_channels_by_team_id(self, team_id, include_private=False):
        """
        List channels of a team.
        Args:
            team_id (str): The team id.
            include_private (bool): If True, uses the beta endpoint
                getAllChannels, which also includes private/shared channels
                the app has access to.
        """
        if include_private:
            url = f"{self.BETA_URL}/teams/{team_id}/channels/getAllChannels"
        else:
            url = f"{self.BASE_URL}/teams/{team_id}/channels"
        return self.__get_all_pages(url)

    def get_channel_by_id(self, team_id, channel_id):
        url = f"{self.BASE_URL}/teams/{team_id}/channels/{channel_id}"
        response = self.http.get(url)
        if response.status_code == 200:
            return response.json()
        raise self.HermesMSGraphError(
            f"Error fetching channel: {response.status_code} - {response.text}"
        )

    # ---------------- Channel Messages ----------------

    def list_channel_messages(self, team_id, channel_id, include_replies=True, progress_callback=None):
        """
        List all messages of a channel (top-level posts).
        Args:
            include_replies (bool): If True, also fetches the replies of
                each top-level message and nests them under the "replies" key.
            progress_callback: optional callable invoked as
                progress_callback(pages_fetched, items_so_far) after each
                page of top-level messages is fetched. Useful for channels
                with a large message history, where a single call can take
                several minutes to paginate through the full history.
        """
        url = f"{self.BASE_URL}/teams/{team_id}/channels/{channel_id}/messages"
        messages = self.__get_all_pages(url, progress_callback=progress_callback)

        if include_replies:
            for message in messages:
                message["replies"] = self.list_channel_message_replies(
                    team_id, channel_id, message["id"]
                )

        return messages

    def list_channel_message_replies(self, team_id, channel_id, message_id, progress_callback=None):
        url = (
            f"{self.BASE_URL}/teams/{team_id}/channels/{channel_id}"
            f"/messages/{message_id}/replies"
        )
        return self.__get_all_pages(url, progress_callback=progress_callback)

    def delta_channel_messages(self, team_id, channel_id, delta_link=None):
        """
        Incremental sync of channel messages using Microsoft Graph delta query.
        Pass the deltaLink returned from a previous call to fetch only changes
        since then; omit it to start a new delta chain.

        Returns:
            dict: {"messages": [...], "delta_link": "<url to store for next sync>"}
        """
        url = delta_link or (
            f"{self.BASE_URL}/teams/{team_id}/channels/{channel_id}/messages/delta"
        )

        messages = []
        next_url = url
        delta_link_result = None

        while next_url:
            response = self.http.get(next_url)
            if response.status_code != 200:
                raise self.HermesMSGraphError(
                    f"Error fetching delta messages: {response.status_code} - {response.text}"
                )
            data = response.json()
            messages.extend(data.get("value", []))
            next_url = data.get("@odata.nextLink")
            if not next_url:
                delta_link_result = data.get("@odata.deltaLink")

        return {"messages": messages, "delta_link": delta_link_result}

    # ---------------- Chats ----------------

    def list_chats_by_user_id(self, user_id, expand_members=True, progress_callback=None):
        """
        List every chat (1:1, group, meeting) a user participates in.
        Args:
            expand_members (bool): If True, expands chat members inline.
            progress_callback: optional callable invoked as
                progress_callback(pages_fetched, items_so_far) after each page.
        """
        url = f"{self.BASE_URL}/users/{user_id}/chats"
        if expand_members:
            url += "?$expand=members"
        return self.__get_all_pages(url, progress_callback=progress_callback)

    def get_chat_by_id(self, chat_id):
        url = f"{self.BASE_URL}/chats/{chat_id}"
        response = self.http.get(url)
        if response.status_code == 200:
            return response.json()
        raise self.HermesMSGraphError(
            f"Error fetching chat: {response.status_code} - {response.text}"
        )

    def list_chat_members(self, chat_id):
        url = f"{self.BASE_URL}/chats/{chat_id}/members"
        return self.__get_all_pages(url)

    def list_chat_messages(self, chat_id, progress_callback=None):
        """
        List all messages of a chat (1:1 or group), oldest pagination handled
        automatically. Requires Chat.Read.All (application permission).
        Args:
            progress_callback: optional callable invoked as
                progress_callback(pages_fetched, items_so_far) after each page.
        """
        url = f"{self.BASE_URL}/chats/{chat_id}/messages"
        return self.__get_all_pages(url, progress_callback=progress_callback)

    def delta_chat_messages(self, chat_id, delta_link=None):
        """
        Incremental sync of chat messages using Microsoft Graph delta query.
        Same usage pattern as delta_channel_messages.
        """
        url = delta_link or f"{self.BASE_URL}/chats/{chat_id}/messages/delta"

        messages = []
        next_url = url
        delta_link_result = None

        while next_url:
            response = self.http.get(next_url)
            if response.status_code != 200:
                raise self.HermesMSGraphError(
                    f"Error fetching delta chat messages: {response.status_code} - {response.text}"
                )
            data = response.json()
            messages.extend(data.get("value", []))
            next_url = data.get("@odata.nextLink")
            if not next_url:
                delta_link_result = data.get("@odata.deltaLink")

        return {"messages": messages, "delta_link": delta_link_result}

    # ---------------- Attachments referenced in messages ----------------

    def download_hosted_content(self, url, file_path):
        """
        Downloads binary hosted content referenced by a message
        (e.g. inline images at .../hostedContents/{id}/$value).
        """
        response = self.http.get(url)
        if response.status_code != 200:
            raise self.HermesMSGraphError(
                f"Error downloading hosted content: {response.status_code} - {response.text}"
            )
        with open(file_path, "wb") as file:
            file.write(response.content)
        return file_path
