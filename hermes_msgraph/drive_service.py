import os
from urllib.parse import urlparse

import requests

from exceptions import HermesMSGraphError


class DriveService:
    """
    Service to list SharePoint sites and download files from SharePoint
    document libraries and OneDrive (application permissions:
    Sites.Read.All, Files.Read.All).
    """

    BASE_URL = "https://graph.microsoft.com/v1.0"

    def __init__(self, http_client):
        self.http = http_client

    def list_sharepoint_sites(self):
        """
        List every root site collection in the tenant.
        """
        url = f"{self.BASE_URL}/sites?$select=id,name,webUrl,siteCollection&$filter=siteCollection/root ne null"
        return self.http.get_json_response_by_url(url, get_value=True)

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

    def __download_drive_item(self, item, local_path):
        download_url = item.get("@microsoft.graph.downloadUrl")
        if not download_url:
            return

        os.makedirs(os.path.dirname(local_path), exist_ok=True)

        response = requests.get(download_url)
        if response.status_code != 200:
            raise HermesMSGraphError(
                f"Error downloading file '{item.get('name')}': {response.status_code} - {response.text}"
            )

        with open(local_path, "wb") as file:
            file.write(response.content)

    def __download_drive_folder(self, drive_id, item_id, local_path):
        """
        Recursively downloads every file under a drive folder, preserving
        the folder structure locally.
        """
        url = f"{self.BASE_URL}/drives/{drive_id}/items/{item_id}/children"
        children = self.http.get_json_response_by_url(url, get_value=True)

        for child in children:
            child_local_path = os.path.join(local_path, child["name"])

            if "folder" in child:
                self.__download_drive_folder(drive_id, child["id"], child_local_path)
            else:
                self.__download_drive_item(child, child_local_path)

    def download_onedrive_files(self, user_email_or_id: str, local_path: str):
        """
        Downloads every file from a user's OneDrive to a local folder,
        preserving the original folder structure.
        :param user_email_or_id: The user's email address or user ID.
        :param local_path: Local directory to save the files into.
        """
        drive_id = self.__get_drive_id_for_user(user_email_or_id)

        url = f"{self.BASE_URL}/drives/{drive_id}/root"
        root = self.http.get_json_response_by_url(url, get_value=False)

        self.__download_drive_folder(drive_id, root["id"], local_path)

    def download_sharepoint_site(self, site: str, local_path: str):
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

        self.__download_drive_folder(drive_id, root["id"], local_path)

    def download_all_sharepoint_sites(self, local_path: str):
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
                self.download_sharepoint_site(site["id"], site_local_path)
            except HermesMSGraphError as e:
                raise HermesMSGraphError(
                    f"Error backing up site '{site_name}' ({site['webUrl']}): {e}"
                ) from e
