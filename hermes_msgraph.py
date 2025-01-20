import json
import os
import sys
import base64
import pandas as pd

sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))
from classes.hermeshttp import HermesHttp

class HermesGraphAPIError(Exception):
    """Custom exception for errors related to HermesGraphAPI."""
    def __init__(self, message, error_code=None):
        super().__init__(message)
        self.error_code = error_code

    pass

class HermesMSGraph:
    """
    Class to interact with the Microsoft Graph API for sending and reading emails.

    The class contains methods to obtain an access token, send emails, read email messages, organize data into a DataFrame, and save it to a JSON file.
    """

    def __init__(self, client_id, client_secret, tenant_id):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.http = HermesHttp(client_id, client_secret, tenant_id)

    def send_email(self, sender_mail, subject, body, to_address, cc_address=None):
        url = f"https://graph.microsoft.com/v1.0/users/{sender_mail}/sendMail"

        payload = {
            "message": {
                "subject": subject,
                "body": {"contentType": "Text", "content": body},
                "toRecipients": [{"emailAddress": {"address": to_address}}],
                "ccRecipients": (
                    [{"emailAddress": {"address": cc_address}}] if cc_address else []
                ),
            },
            "saveToSentItems": "true",
        }

        response = self.http.post(url, payload=payload)

        if response.status_code == 200:
            print("Email sent successfully!")
        else:
            print(f"Error sending email: {response.status_code}")


    def __get_json_response_by_url(self, url, get_value=True):
        response = self.http.get(url)
    
        if response.status_code == 200:
            try:
                json_response = response.json()
                return json_response.get("value", []) if get_value else json_response
            except ValueError:
                print(f"Invalid JSON response from {url}")
                return None
        else:
            print(f"HTTP Error: {response.status_code} - {response.text}")
            return None


    def list_email_attachments(self, mailbox_address, message_id):
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox_address}/messages/{message_id}/attachments"
        data_json = self.__get_json_response_by_url(url, get_value=True)

        return data_json

    def download_attachment(
        self, mailbox_address, message_id, attachment_id, file_name
    ):
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox_address}/messages/{message_id}/attachments/{attachment_id}"
        data_json = self.__get_json_response_by_url(url, get_value=False)

        if data_json:
            directory = os.path.dirname(file_name)
            os.makedirs(directory, exist_ok=True)

            with open(file_name, "wb") as file:
                file.write(base64.b64decode(data_json["contentBytes"]))
            print(f"Attachment saved as {file_name}")

    def list_mailbox_folders(self, mailbox_address):
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox_address}/mailfolders/delta?$select=displayname"
        
        all_folders = []
        while url:
            data = self.__get_json_response_by_url(url, get_value= False)

            if data:
                #mail_folders = response.json().get("value", [])
                all_folders.extend(data.get("value", []))

                url = data.get("@odata.nextLink")
            else:
                print(f"No folders found for {mailbox_address}")

        return all_folders

    def get_mailbox_folders(self, mailbox_address):
        mail_folders = self.list_mailbox_folders(mailbox_address)

        df_mail_folders = pd.DataFrame(mail_folders)
        if not df_mail_folders.empty:
            return df_mail_folders


    def __build_email_query_params(
        self,
        subject=None,
        sender=None,
        n_of_messages=None,
        has_attachments=None,
        greater_than_date=None,
        less_than_date=None,
    ):
        filters = []
        query_params = []

        def add_filter(condition):
            if condition:
                filters.append(condition)

        add_filter(f"sender/emailAddress/address eq '{sender}'" if sender else None)
        add_filter(self.__url_filter_subject(subject) if subject else None)
        add_filter(f"receivedDateTime gt {greater_than_date}" if greater_than_date else None)
        add_filter(f"receivedDateTime lt {less_than_date}" if less_than_date else None)
        add_filter(f"hasAttachments eq {str(has_attachments).lower()}" if has_attachments else None)

        if filters:
            query_params.append(f"$filter={' and '.join(filters)}")

        if n_of_messages:
            query_params.append(f"$top={n_of_messages}")

        return "&".join(query_params) if query_params else ""

    def __get_folder_id(self, mailbox_address, folder_name):
        df_folders = self.get_mailbox_folders(mailbox_address)
        if folder_name and not df_folders.empty:
            folder_row = df_folders.loc[df_folders["displayName"] == folder_name]
            if not folder_row.empty:
                return folder_row["id"].iloc[0]
        raise ValueError(f"Folder '{folder_name}' not found for {mailbox_address}")

    def __read_emails(
        self,
        mailbox_address,
        subject,
        folder,
        sender,
        n_of_messages,
        has_attachments,
        messages_json_path,
        greater_than_date,
        less_than_date,
    ):
        
        folder_id = self.__get_folder_id(mailbox_address, folder) if folder else None
        folder_path = f"/mailFolders/{folder_id}" if folder_id else ""


        email_filter_url = self.__build_email_query_params(
            subject, sender, n_of_messages, has_attachments, greater_than_date, less_than_date
        )
        
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox_address}{folder_path}/messages?{email_filter_url}"
        #print(url)
        data_json = self.__get_json_response_by_url(url, get_value=True)

        if not data_json:
            print(f"No emails found for {mailbox_address}")
            return None
        else:
            if messages_json_path:
                try:
                    with open(messages_json_path, "w+", encoding="utf-8") as file:
                        json.dump(data_json, file, ensure_ascii=False, indent=4)
                except Exception as e:
                    raise RuntimeError(
                        "Unable to save messages to messages.json file"
                    ) from e
            return data_json

    def __json_to_dataframe(self, json_file_path):
        with open(json_file_path, "r") as file:
            json_data = json.load(file)
        df_emails = pd.json_normalize(json_data)
        return df_emails

    def __filter_columns_df_emails(self, df_emails):
        columns = [
            "subject",
            "isRead",
            "sentDateTime",
            "receivedDateTime",
            "sender.emailAddress.name",
            "sender.emailAddress.address",
            "from.emailAddress.name",
            "from.emailAddress.address",
            "bodyPreview",
            "body.contentType",
            "body.content",
            "id",
        ]
        df_emails = df_emails[columns]
        return df_emails

    # Usage example:
    # api = MSGraphAPI()
    # df_emails = api.get_df_emails("anakin.skywalker@github.com")

    def get_emails(
        self,
        mailbox_address,
        subject="",
        folder="",
        sender="",
        n_of_messages=10,
        has_attachments="",
        messages_json_path=None,
        greater_than_date=None,
        less_than_date=None,
        format="dataframe",
        data="all", #simple or all
    ):
        json_emails = self.__read_emails(
            mailbox_address=mailbox_address,
            subject=subject,
            folder=folder,
            sender=sender,
            n_of_messages=n_of_messages,
            has_attachments=has_attachments,
            messages_json_path=messages_json_path,
            greater_than_date=greater_than_date,
            less_than_date=less_than_date,
        )

        if not json_emails:
            return pd.DataFrame()

        df_emails = pd.json_normalize(json_emails)
       
        if data == "simple":
            df_emails = self.__filter_columns_df_emails(df_emails)

        return df_emails.to_json(orient="records") if format == "json" else df_emails

    def __read_email_by_id(self, email_id, mailbox_address):
    
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox_address}/messages/{email_id}"

        response_json = self.__get_json_response_by_url(url, get_value=False)
        if not response_json:
            raise HermesGraphAPIError(f"Failed to retrieve email with ID {email_id}")
        
        return response_json


    def get_email_by_id(self, email_id, mailbox_address):
        email_json = self.__read_email_by_id(email_id, mailbox_address)
        email_df = pd.DataFrame(email_json)
        return email_df

    #legacy
    def get_df_emails(
        self,
        mailbox_address,
        subject="",
        folder="",
        sender="",
        n_of_messages=10,
        has_attachments="",
        messages_json_path=None,
        greater_than_date=None,
        less_than_date=None,
    ):
        emails = self.get_emails(
            mailbox_address=mailbox_address,
            subject=subject,
            folder=folder,
            sender=sender,
            n_of_messages=n_of_messages,
            has_attachments=has_attachments,
            messages_json_path=messages_json_path,
            greater_than_date=greater_than_date,
            less_than_date=less_than_date,
            format="dataframe",
            data="simple",
        )
        return emails

    def get_raw_df_emails(
        self,
        mailbox_address,
        subject="",
        folder="",
        sender="",
        n_of_messages=10,
        has_attachments="",
        messages_json_path=None,
        greater_than_date=None,
        less_than_date=None,
    ):
        emails = self.get_emails(
            mailbox_address=mailbox_address,
            subject=subject,
            folder=folder,
            sender=sender,
            n_of_messages=n_of_messages,
            has_attachments=has_attachments,
            messages_json_path=messages_json_path,
            greater_than_date=greater_than_date,
            less_than_date=less_than_date,
            format="dataframe",
            data="all",
        )
        return emails
