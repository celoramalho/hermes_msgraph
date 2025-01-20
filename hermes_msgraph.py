import json
import requests
import yaml
import os
import base64
import pandas as pd

from classes.hermeshttp import HermesHttp

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
        #self.access_token = self.__get_access_token()

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

    def __ulr_filter_subject(self, subject):
        if subject:
            # Case: starts and ends with asterisk -> "contains"
            if subject.startswith("*") and subject.endswith("*"):
                subject = subject.strip("*")
                subject_filter_url = f"contains(subject, '{subject}')"

            # Case: starts with asterisk -> "ends with"
            elif subject.startswith("*") and not subject.endswith("*"):
                subject = subject.strip("*")
                subject_filter_url = f"startswith(subject, '{subject}')"

            # Case: ends with asterisk -> "starts with"
            elif subject.endswith("*") and not subject.startswith("*"):
                subject = subject.strip("*")
                subject_filter_url = f"endswith(subject, '{subject}')"

            # Case: exact match
            else:
                subject_filter_url = f"subject eq '{subject}'"

        return subject_filter_url

    def list_email_attachments(self, mailbox_address, message_id):
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox_address}/messages/{message_id}/attachments"
        data_json = self.__get_json_response_by_url(url, get_value=True)

        return data_json

    def download_attachment(
        self, mailbox_address, message_id, attachment_id, file_name
    ):
        # URL para acessar o anexo
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox_address}/messages/{message_id}/attachments/{attachment_id}"
        data_json = self.__get_json_response_by_url(url, get_value=False)

        if data_json:
            # Criar diretório se não existir
            directory = os.path.dirname(file_name)
            os.makedirs(directory, exist_ok=True)

            # Salvar o conteúdo do anexo no arquivo
            with open(file_name, "wb") as file:
                file.write(base64.b64decode(data_json["contentBytes"]))
                # file.write(bytes(attachment_data['contentBytes'], 'utf-8'))
            print(f"Attachment saved as {file_name}")

    def list_mailbox_folders(self, mailbox_address):
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox_address}/mailfolders/delta?$select=displayname"
        
        all_folders = []
        # https://stackoverflow.com/questions/42901755/microsoft-graph-outlook-mail-list-all-mail-folders-not-just-the-top-level-o
        # url = f"https://graph.microsoft.com/v1.0/users/{mailbox_address}/messages?$top={n_of_messages}"

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


    def __define_query_utl_by_email_filters(self, mailbox_address, subject, folder, sender, n_of_messages, has_attachments, messages_json_path, greater_than_date, less_than_date):
        email_filter = []
        email_filter_url = []

        if sender:
            email_filter.append(f"sender/emailAddress/address eq '{sender}'")

        if subject:
            email_filter.append(self.__ulr_filter_subject(subject))
            # email_filter.append(f"subject eq '{subject}'")

        if greater_than_date:
            email_filter.append(f"receivedDateTime gt {greater_than_date}")

        if less_than_date:
            email_filter.append(f"receivedDateTime lt {less_than_date}")

        if has_attachments:
            email_filter.append(
                f"hasAttachments eq {str(has_attachments).lower()}"
            )

        if email_filter:
            email_filter_joined = " and ".join(email_filter)
            email_filter_url.append(f"$filter={email_filter_joined}")

        if n_of_messages:
            email_filter_url.append(f"$top={n_of_messages}")

        if email_filter_url:
            email_filter_url = "&".join(email_filter_url)
        else:
            email_filter_url = ""

        return email_filter_url

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
        
        if folder:
            df_folders = self.get_mailbox_folders(mailbox_address)
            folder_id = None
            for index, existing_folder in df_folders.iterrows():
                if existing_folder["displayName"] == folder:
                    folder_id = df_folders.loc[df_folders["displayName"] == folder, "id"].iloc[0]
            
            if not folder_id:
                raise ValueError(f"Folder '{folder}' not found for user '{mailbox_address}'")
            folder = f"/mailFolders/{folder_id}"


        email_filter_url = self.__define_query_utl_by_email_filters(mailbox_address, subject, folder, sender, n_of_messages, has_attachments, messages_json_path, greater_than_date, less_than_date)
        
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox_address}{folder}/messages?{email_filter_url}"
        #print(url)
        data_json = self.__get_json_response_by_url(url, get_value=True)

        if data_json:
            try:
                with open(messages_json_path, "w+", encoding="utf-8") as file:
                    json.dump(data_json, file, ensure_ascii=False, indent=4)
            except Exception as e:
                raise RuntimeError(
                    "Unable to save messages to messages.json file"
                ) from e
        else:
            print(f"No emails found for {mailbox_address}")

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
        messages_json_path="messages.json",
        greater_than_date=None,
        less_than_date=None,
        format="dataframe",
        data="all", #simple or all
    ):
        """
        Public method to retrieve email messages and organize them into a DataFrame with all received attributes.

        Parameters:
        ----------
        mailbox_address : str
            Email address from which to retrieve messages.
        subject : str, optional
            Subject to filter by, default is an empty string.
        folder : str, optional
            Folder to retrieve messages from, default is an empty string.
        sender : str, optional
            Sender's email address to filter by, default is an empty string.
        n_of_messages : int, optional
            Number of emails to retrieve, default is 10.
        subject_filter : str, optional
            Filter to apply to email subjects, default is an empty string.
        messages_json_path : str, optional
            Path to the JSON file where messages will be saved, default is 'messages.json'.

        Returns:
        -------
        df_emails : pandas.DataFrame
            Full DataFrame with email messages and all attributes.
        """
        self.__read_emails(
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
        df_raw_emails = self.__json_to_dataframe(messages_json_path)
        #if subject_filter != "":
        #    df_raw_emails = self.__filter_subject(df_raw_emails, subject_filter)
       
        if data == "simple":
            df_emails = self.__filter_columns_df_emails(df_raw_emails)
        elif data == "all":
            # print("No emails found using the parameters passed.")
            df_emails = df_raw_emails

            
        if format == "json":
            json_emails = df_emails.to_json(orient="records")
            emails = json.loads(json_emails)
        elif format == "dataframe":
            emails = df_emails

        return emails

    def __get_json_response_by_url(self, url, get_value = True):
        response = self.http.get(url)

        if response.status_code == 200:
            json_response = response.json().get("value", [])

            if json_response == None or not get_value:
                json_response = response.json()

        else:
            print(f"Error in HTTP request to {url}: {response.status_code} - {response.text}")
            json_response = None

        return json_response

    def __read_email_by_id(self, email_id, mailbox_address, messages_json_path="messages.json"):
    
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox_address}/messages/{email_id}"

        response_json = self.__get_json_response_by_url(url)

        email_json = response_json
        
        return email_json


    def get_email_by_id(self, email_id, mailbox_address, messages_json_path="messages.json"):
        """
        Public method to retrieve email messages by their ID.

        Parameters:
        ----------
        email_id : str
            ID of the email to retrieve.
        messages_json_path : str, optional
            Path to the JSON file where messages will be saved, default is 'messages.json'.

        Returns:
        -------
        df_emails : pandas.DataFrame
            DataFrame with email messages.
        """
        email_json = self.__read_email_by_id(email_id, mailbox_address, messages_json_path)
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
        messages_json_path="messages.json",
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
        messages_json_path="messages.json",
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
