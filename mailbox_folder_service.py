from typing import List, Dict, Union
import pandas as pd
from http_client import HttpClient
from exceptions import HermesMSGraphError
from typing import Optional


class MailboxFolderService:
    def __init__(self, http_client: HttpClient):
        self.http = http_client

    def get_mailbox_folders(self, mailbox_address: str) -> pd.DataFrame:
        """
        Retrieve all mailbox folders for a given mailbox address and return as a DataFrame.
        :param mailbox_address: The email address of the mailbox.
        :return: DataFrame containing mailbox folders.
        """
        mail_folders = self.list_mailbox_folders(mailbox_address)
        df_mail_folders = pd.DataFrame(mail_folders)
        return df_mail_folders if not df_mail_folders.empty else pd.DataFrame()

    def list_mailbox_folders(self, mailbox_address: str) -> List[Dict]:
        """
        List all mailbox folders for a given mailbox address.
        :param mailbox_address: The email address of the mailbox.
        :return: List of mailbox folders.
        """
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox_address}/mailfolders/delta?$select=displayname"
        all_folders = []

        while url:
            data = self._fetch_data(url)
            if data:
                all_folders.extend(data.get("value", []))
                url = data.get("@odata.nextLink")
            else:
                raise HermesMSGraphError(f"Failed to retrieve folders for {mailbox_address}")

        return all_folders

    def validate_folder_id(self, mailbox_address: str, folder_id: str) -> bool:
        """
        Validate if a folder ID exists in the mailbox.
        :param mailbox_address: The email address of the mailbox.
        :param folder_id: The ID of the folder to validate.
        :return: True if the folder ID exists, False otherwise.
        """
        df_folders = self.get_mailbox_folders(mailbox_address)
        if not df_folders.empty:
            return not df_folders.loc[df_folders["id"] == folder_id].empty
        return False

    def get_folder_id(self, mailbox_address: str, folder_name: str) -> Optional[str]:
        """
        Retrieve the folder ID for a given folder name.
        :param mailbox_address: The email address of the mailbox.
        :param folder_name: The name of the folder.
        :return: The folder ID if found, None otherwise.
        """
        df_folders = self.get_mailbox_folders(mailbox_address)

        if folder_name and not df_folders.empty:
            folder_row = df_folders.loc[df_folders["displayName"] == folder_name]
            if not folder_row.empty:
                return folder_row["id"].iloc[0]
        return None

    def _fetch_data(self, url: str) -> Dict:
        """
        Helper method to fetch data from a given URL with error handling.
        :param url: The URL to fetch data from.
        :return: JSON response as a dictionary.
        """
        response = self.http.get(url)
        if response.status_code != 200:
            raise HermesMSGraphError(f"Failed to fetch data: {response.status_code} - {response.text}")
        return response.json()