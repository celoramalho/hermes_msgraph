import json
import os
from urllib.parse import quote, unquote
import sys
import base64
import pandas as pd

sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))
from classes.hermeshttp import HermesHttp

class HermesMSGraph:
    """
    Class to interact with the Microsoft Graph API for sending and reading emails.

    The class contains methods to obtain an access token, send emails, read email messages, organize data into a DataFrame, and save it to a JSON file.
    """


    class HermesMSGraphError(Exception):
        """Custom exception for errors related to HermesGraphAPI."""
        def __init__(self, message, error_code=None):
            super().__init__(message)
            self.error_code = error_code

        pass


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

        if response.status_code == 200 or response.status_code == 202:
            print("Email sent successfully!")
        else:
            raise self.HermesMSGraphError(f"Error sending email: {response.status_code} - {response.text}")


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
            raise self.HermesMSGraphError(f"Error fetching data from {url}: {response.status_code} - {response.text}")


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
                raise self.HermesMSGraphError(f"Failed to retrieve folders for {mailbox_address}")

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

    
    def __validate_folder_id(self, mailbox_address, folder_id):
        df_folders = self.get_mailbox_folders(mailbox_address)
        
        if not df_folders.empty:
            folder_row = df_folders.loc[df_folders["id"] == folder_id]
            if not folder_row.empty:
                return True
        return False
    
    def __get_folder_id(self, mailbox_address, folder_name):
        df_folders = self.get_mailbox_folders(mailbox_address)
        
        if folder_name and not df_folders.empty:
            folder_row = df_folders.loc[df_folders["displayName"] == folder_name]
            if not folder_row.empty:
                return folder_row["id"].iloc[0]
        raise self.HermesMSGraphError(f"Folder '{folder_name}' not found for {mailbox_address}")
    
    def __get_folder_name_by_id(self, mailbox_address, folder_id):
        df_folders = self.get_mailbox_folders(mailbox_address)
        
        if not df_folders.empty:
            folder_row = df_folders.loc[df_folders["id"] == folder_id]
            if not folder_row.empty:
                return folder_row["displayName"].iloc[0]
        raise self.HermesMSGraphError(f"Folder with ID {folder_id} not found for {mailbox_address}")
    
    
    def __verify_if_str_is_encoded(self, string):
        """
        Verifies if a string is already URL-encoded.

        Args:
            string (str): The string to verify.

        Returns:
            bool: True if the string is already encoded, False otherwise.
        """
        return string == quote(unquote(string))

    def __encode_str_to_url(self, string):
        """
        Encodes a string to be safely used in a URL.

        Args:
            string (str): The string to encode.

        Returns:
            str: The URL-encoded string.
        """
        if self.__verify_if_str_is_encoded(string):
            return string
        
        return quote(string)
    
    def __url_filter_subject(self, subject):
        """
        Builds a filter for the subject in Microsoft Graph API query parameters.

        Args:
            subject (str): The subject string to filter, supports wildcards '*' at the beginning or end.

        Returns:
            str: The formatted filter string for the subject or an empty string if subject is None or empty.
        """
        if subject:
            # Case: starts and ends with asterisk -> "contains"
            if subject.startswith("*") and subject.endswith("*"):
                subject = subject.strip("*")
                subject_filter_url = f"contains(subject, '{subject}')"

            # Case: starts with asterisk -> "ends with"
            elif subject.startswith("*"):
                subject = subject.strip("*")
                subject_filter_url = f"endswith(subject, '{subject}')"

            # Case: ends with asterisk -> "starts with"
            elif subject.endswith("*"):
                raise self.HermesMSGraphError("Microsoft Graph API does not support endswith filter for subject")
                subject = subject.strip("*")
                subject_filter_url = f"startswith(subject, '{subject}')"

            # Case: exact match
            else:
                subject_filter_url = f"subject eq '{subject}'"

            return subject_filter_url

        # Default return for invalid or empty subject
        return ""

    
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
        
        self.__validate_parameters(
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


        if folder:
            folder_id = self.__get_folder_id(mailbox_address, folder)
            folder_path = f"/mailFolders/{folder_id}" 
        else:
            folder_path = ""


        filter_suffix = self.__build_email_query_params(
            subject, sender, n_of_messages, has_attachments, greater_than_date, less_than_date
        )
        
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox_address}{folder_path}/messages?{filter_suffix}"
        

        data_json = self.__get_json_response_by_url(url, get_value=True)
        
        if messages_json_path:
            try:
                with open(messages_json_path, "w+", encoding="utf-8") as file:
                    json.dump(data_json, file, ensure_ascii=False, indent=4)
            except Exception as e:
                raise self.HermesMSGraphError(
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
        subject=None,
        folder=None,
        sender=None,
        n_of_messages=10,
        has_attachments= "",
        messages_json_path=None,
        greater_than_date=None,
        less_than_date=None,
        format=list,
        data="all", #simple or all
    ):
        
        valid_formats = [list, pd.DataFrame]
        
        if format not in valid_formats:
            raise self.HermesMSGraphError("Invalid format. Must be 'dataframe' or 'list'")

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

        return  pd.json_normalize(json_emails) if format == pd.DataFrame else json_emails


    def __read_email_by_id(self, email_id, mailbox_address):
        
        #email_id = self.__encode_str_to_url(email_id)
    
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox_address}/messages/{email_id}"

        response_json = self.__get_json_response_by_url(url, get_value=False)
        if not response_json:
            raise self.HermesMSGraphError(f"Failed to retrieve email with ID {email_id}")
        
        return response_json
    
    def move_email_to_folder(self, email_id, mailbox_address, folder_name=None, folder_id=None):
        email_id = self.__encode_str_to_url(email_id)
        
        if folder_name:
            folder_id = self.__get_folder_id(mailbox_address, folder_name)
            folder_id_encoded = self.__encode_str_to_url(folder_id)
            
        elif folder_id:
            folder_id_encoded = self.__encode_str_to_url(folder_id)
            folder_name = self.__get_folder_name_by_id(mailbox_address, folder_id)
        
        else:
            raise self.HermesMSGraphError("Either folder_id or folder_name must be provided")
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox_address}/messages/{email_id}/move"
        
        payload = {
            "destinationId": folder_id
        }
        
        #print(url)
        #print(payload)
        
        response = self.http.post(url, payload=payload)
        #print(response.status_code)
        if response.status_code == 403:
            raise self.HermesMSGraphError(f"Error moving email to folder: {response.status_code} - {response.text}")

    def get_email_by_id(self, email_id, mailbox_address):
        email_json = self.__read_email_by_id(email_id, mailbox_address)
        #print(email_json)
        email_df = pd.json_normalize(email_json)
        return email_json

    def foward_email_by_id(self, email_id, mailbox_address, to_address, comment=None):
        """
        Forwards an email by its ID to a specified recipient.

        Args:
            email_id (str): The ID of the email to forward.
            mailbox_address (str): The email address of the mailbox.
            to_address (str): The recipient's email address.
            comment (str, optional): A comment to include with the forwarded email. Defaults to None.
        """
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox_address}/messages/{email_id}/forward"

        payload = {
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": to_address
                    }
                }
            ]
        }

        if comment:
            payload["comment"] = comment

    # Send the POST request
        self.http.post(url, payload=payload)

    
    def list_msgraph_permissions(self):
        self.http.list_msgraph_permisions()

    def get_user_id_by_email(self, email_address):
        url = f"https://graph.microsoft.com/v1.0/users/{email_address}"
        
        response = self.http.get(url)
        user_data = response.json()

        if response.status_code == 200:
            return user_data.get("id")
        else:
            print(f"Error fetching user ID: {response.status_code} - {response.text}")
            return None
        
    def list_plans_by_group_id(self, group_id, data="all"):
        url = f"https://graph.microsoft.com/v1.0/groups/{group_id}/planner/plans"
        response = self.http.get(url)

        plans = response.json().get("value", [])

        planner_data = []
        for plan in plans:
            planner_data.append({
                "id": plan["id"],
                "title": plan["title"],
                "owner": plan["owner"]
            })

        return plans if data == "all" else planner_data
    
    def list_visible_plans_by_user_id(self, user_id):
        url = f"https://graph.microsoft.com/v1.0/users/{user_id}/planner/plans"
        response = self.http.get(url)

        print(response.json())

        plans = response.json().get("value", [])

        planner_data = []
        for plan in plans:
            planner_data.append({
                "id": plan["id"],
                "title": plan["title"],
                "owner": plan["owner"]["user"]["displayName"]
            })

        return planner_data
    

    def list_tasks_by_user_id(self, user_id):
        url = f"https://graph.microsoft.com/v1.0/users/{user_id}/planner/tasks"
        response = self.http.get(url)

        print(response.json())

        tasks = response.json().get("value", [])

        return tasks


    def list_tasks_by_plan_id(self, planner_id):
        url = f"https://graph.microsoft.com/v1.0/planner/plans/{planner_id}/tasks"
        response = self.http.get(url)
        if response.status_code != 200:
            raise self.HermesMSGraphError(f"Error fetching tasks: {response.status_code} - {response.text}")
        
        tasks = response.json().get("value", [])
        detailed_tasks = []

        for task in tasks:
            task_id = task["id"]
            details_url = f"https://graph.microsoft.com/v1.0/planner/tasks/{task_id}/details"
            details_response = self.http.get(details_url)
            if details_response.status_code == 200:
                task_details = details_response.json()
                task["body"] = task_details.get("description", "N/A")  # Add the task body/description
            else:
                task["body"] = "N/A"  # Handle cases where details cannot be fetched
            detailed_tasks.append(task)

        return detailed_tasks


    def __validate_parameters(
        self,
        mailbox_address,
        subject=None,
        folder=None,
        sender=None,
        n_of_messages=None,
        has_attachments= "",
        messages_json_path=None,
        greater_than_date=None,
        less_than_date=None,
    ):
        if not isinstance(mailbox_address, str) or not mailbox_address:
            raise self.HermesMSGraphError("Invalid mailbox_address. Must be a non-empty string.")
        
        if subject is not None and not isinstance(subject, str):
            raise self.HermesMSGraphError("Invalid subject. Must be a string.")
        
        if folder is not None and not isinstance(folder, str):
            raise self.HermesMSGraphError("Invalid folder. Must be a string.")
        
        if sender is not None and not isinstance(sender, str):
            raise self.HermesMSGraphError("Invalid sender. Must be a string.")
        
        if n_of_messages != "all" and (not isinstance(n_of_messages, int) or n_of_messages <= 0):
            raise self.HermesMSGraphError("Invalid n_of_messages. Must be a positive integer or 'all'.")
        
        if has_attachments not in ["", True, False]:
            raise self.HermesMSGraphError("Invalid has_attachments. Must be True, False, or an empty string.")
        
        if messages_json_path is not None and not isinstance(messages_json_path, str):
            raise self.HermesMSGraphError("Invalid messages_json_path. Must be a string.")
        
        if greater_than_date is not None and not isinstance(greater_than_date, str):
            raise self.HermesMSGraphError("Invalid greater_than_date. Must be a string in ISO format.")
        
        if less_than_date is not None and not isinstance(less_than_date, str):
            raise self.HermesMSGraphError("Invalid less_than_date. Must be a string in ISO format.")