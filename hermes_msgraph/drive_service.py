import os
import time
from urllib.parse import urlparse

import requests

from exceptions import HermesMSGraphError


class DriveService:
    """
    Service to list SharePoint sites/drives and download files from
    SharePoint document libraries and OneDrive (application permissions:
    Sites.Read.All, Files.Read.All).
    """

    BASE_URL = "https://graph.microsoft.com/v1.0"

    def __init__(self, http_client):
        self.http = http_client

    def __get_all_pages(self, url, max_retries=5, progress_callback=None):
        """
        Follows @odata.nextLink until every page has been collected,
        retrying on 429 (throttling) using the Retry-After header, and on
        timeouts/connection errors with a short exponential backoff.
        Mirrors TeamsService.__get_all_pages so every paginated endpoint in
        the library behaves the same way under throttling/network errors.
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
                    raise HermesMSGraphError(
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
                    raise HermesMSGraphError(
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
                raise HermesMSGraphError(
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

    # ---------------- Sites / Drives ----------------

    def list_sharepoint_sites(self):
        """
        List every root site collection in the tenant.
        """
        url = f"{self.BASE_URL}/sites?$select=id,name,webUrl,siteCollection&$filter=siteCollection/root ne null"
        return self.__get_all_pages(url)

    def list_drives_by_site_id(self, site_id):
        """
        List every document library (drive) of a SharePoint site. A site
        can have more than one library beyond the default "Documents".
        """
        url = f"{self.BASE_URL}/sites/{site_id}/drives"
        return self.__get_all_pages(url)

    def list_drive_children(self, drive_id, item_id="root", progress_callback=None):
        """
        List the direct children (files and folders) of a drive item,
        paginating through every page via @odata.nextLink.
        """
        url = f"{self.BASE_URL}/drives/{drive_id}/items/{item_id}/children"
        return self.__get_all_pages(url, progress_callback=progress_callback)

    def get_drive_item_by_path(self, drive_id, item_path):
        """
        Fetch a single drive item (file or folder) by its path relative to
        the drive root, e.g. "Folder/Subfolder/file.xlsx".
        """
        item_path = item_path.strip("/")
        url = f"{self.BASE_URL}/drives/{drive_id}/root:/{item_path}"
        response = self.http.get(url)
        if response.status_code == 200:
            return response.json()
        raise HermesMSGraphError(
            f"Error fetching drive item '{item_path}': {response.status_code} - {response.text}"
        )

    def delta_drive_items(self, drive_id, delta_link=None):
        """
        Incremental sync of a drive's items using Microsoft Graph delta
        query. Pass the deltaLink returned from a previous call to fetch
        only changes (new/updated/deleted items) since then; omit it to
        start a new delta chain from the drive root.

        Returns:
            dict: {"items": [...], "delta_link": "<url to store for next sync>"}
        """
        url = delta_link or f"{self.BASE_URL}/drives/{drive_id}/root/delta"

        items = []
        next_url = url
        delta_link_result = None

        while next_url:
            response = self.http.get(next_url)
            if response.status_code != 200:
                raise HermesMSGraphError(
                    f"Error fetching delta drive items: {response.status_code} - {response.text}"
                )
            data = response.json()
            items.extend(data.get("value", []))
            next_url = data.get("@odata.nextLink")
            if not next_url:
                delta_link_result = data.get("@odata.deltaLink")

        return {"items": items, "delta_link": delta_link_result}

    def __resolve_site_id(self, site):
        """
        Accepts either a site id (already in the "hostname,siteId,webId" or
        plain GUID form) or a SharePoint site URL, and returns a site id
        usable in /sites/{site-id} calls.
        """
        if not site:
            raise HermesMSGraphError("Invalid site. Must be a site ID or a SharePoint site URL.")

        if not site.lower().startswith("http"):
            return site

        parsed = urlparse(site)
        hostname = parsed.netloc
        server_relative_path = parsed.path.rstrip("/")

        url = f"{self.BASE_URL}/sites/{hostname}:{server_relative_path}"
        response = self.http.get(url)

        if response.status_code == 200:
            return response.json()["id"]

        raise HermesMSGraphError(
            f"Error resolving site '{site}': {response.status_code} - {response.text}"
        )

    def __get_drive_id_for_site(self, site_id):
        url = f"{self.BASE_URL}/sites/{site_id}/drive"
        response = self.http.get(url)

        if response.status_code == 200:
            return response.json()["id"]

        raise HermesMSGraphError(
            f"Error fetching default drive for site '{site_id}': {response.status_code} - {response.text}"
        )

    def __get_drive_id_for_user(self, user_email_or_id):
        url = f"{self.BASE_URL}/users/{user_email_or_id}/drive"
        response = self.http.get(url)

        if response.status_code == 200:
            return response.json()["id"]

        raise HermesMSGraphError(
            f"Error fetching OneDrive for user '{user_email_or_id}': {response.status_code} - {response.text}"
        )

    # ---------------- Downloads ----------------

    def __download_drive_item(self, item, local_path, max_retries=5, chunk_size=1024 * 1024):
        download_url = item.get("@microsoft.graph.downloadUrl")
        if not download_url:
            return

        os.makedirs(os.path.dirname(local_path), exist_ok=True)

        retries = 0
        network_retries = 0
        max_network_retries = 5

        while True:
            try:
                response = self.http.session.get(
                    download_url, timeout=self.http.timeout, stream=True
                )
            except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as erro:
                if network_retries >= max_network_retries:
                    raise HermesMSGraphError(
                        f"Too many network retries downloading '{item.get('name')}': {erro}"
                    )
                wait_seconds = 2 ** network_retries
                print(
                    f"[hermes_msgraph] Network error downloading '{item.get('name')}' ({erro}). "
                    f"Retrying in {wait_seconds}s (attempt {network_retries + 1}/{max_network_retries})..."
                )
                time.sleep(wait_seconds)
                network_retries += 1
                continue

            if response.status_code == 429:
                if retries >= max_retries:
                    raise HermesMSGraphError(
                        f"Too many retries after 429 throttling downloading '{item.get('name')}'"
                    )
                retry_after = int(response.headers.get("Retry-After", "5"))
                print(
                    f"[hermes_msgraph] Throttled (429) downloading '{item.get('name')}'. "
                    f"Retrying in {retry_after}s (attempt {retries + 1}/{max_retries})..."
                )
                time.sleep(retry_after)
                retries += 1
                continue

            if response.status_code != 200:
                raise HermesMSGraphError(
                    f"Error downloading file '{item.get('name')}': {response.status_code} - {response.text}"
                )

            tmp_path = f"{local_path}.part"
            with open(tmp_path, "wb") as file:
                for chunk in response.iter_content(chunk_size=chunk_size):
                    if chunk:
                        file.write(chunk)
            os.replace(tmp_path, local_path)
            return

    def __download_drive_folder(self, drive_id, item_id, local_path, progress_callback=None):
        """
        Recursively downloads every file under a drive folder, preserving
        the folder structure locally. Paginates through every page of
        children so folders with more items than a single page (commonly
        200) aren't silently truncated.
        """
        children = self.list_drive_children(drive_id, item_id, progress_callback=progress_callback)

        for child in children:
            child_local_path = os.path.join(local_path, child["name"])

            if "folder" in child:
                self.__download_drive_folder(
                    drive_id, child["id"], child_local_path, progress_callback=progress_callback
                )
            else:
                self.__download_drive_item(child, child_local_path)

    def download_onedrive_files(self, user_email_or_id: str, local_path: str, progress_callback=None):
        """
        Downloads every file from a user's OneDrive to a local folder,
        preserving the original folder structure.
        :param user_email_or_id: The user's email address or user ID.
        :param local_path: Local directory to save the files into.
        """
        drive_id = self.__get_drive_id_for_user(user_email_or_id)

        url = f"{self.BASE_URL}/drives/{drive_id}/root"
        root = self.http.get_json_response_by_url(url, get_value=False)

        self.__download_drive_folder(drive_id, root["id"], local_path, progress_callback=progress_callback)

    def download_sharepoint_site(self, site: str, local_path: str, progress_callback=None):
        """
        Downloads every file from a SharePoint site's default document
        library to a local folder, preserving the original folder structure.
        :param site: The site ID, or the SharePoint site URL
            (e.g. "https://contoso.sharepoint.com/sites/Marketing").
        :param local_path: Local directory to save the files into.
        """
        site_id = self.__resolve_site_id(site)
        drive_id = self.__get_drive_id_for_site(site_id)

        url = f"{self.BASE_URL}/drives/{drive_id}/root"
        root = self.http.get_json_response_by_url(url, get_value=False)

        self.__download_drive_folder(drive_id, root["id"], local_path, progress_callback=progress_callback)

    def download_all_sharepoint_sites(self, local_path: str, progress_callback=None):
        """
        Full backup: lists every SharePoint site in the tenant and downloads
        each one into its own subfolder (named after the site) under
        local_path, preserving each site's folder structure.
        :param local_path: Local base directory. One subfolder per site is
            created inside it.
        """
        sites = self.list_sharepoint_sites()

        for site in sites:
            site_name = site.get("name") or site["id"]
            site_local_path = os.path.join(local_path, site_name)

            try:
                self.download_sharepoint_site(site["id"], site_local_path, progress_callback=progress_callback)
            except HermesMSGraphError as e:
                raise HermesMSGraphError(
                    f"Error backing up site '{site_name}' ({site['webUrl']}): {e}"
                ) from e
