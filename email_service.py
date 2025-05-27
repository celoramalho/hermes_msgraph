import pandas as pd
from http_client import HttpClient
from exceptions import HermesMSGraphError
from mailbox_folder_service import MailboxFolderService


class EmailService:
    def __init__(self, http_client: HttpClient):
        self.http = http_client
        self.HermesMSGraphError = HermesMSGraphError
        self.MailboxFolderService = MailboxFolderService(http_client)

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
        

    def list_email_attachments(self, mailbox_address, email_id):
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox_address}/messages/{email_id}/attachments"
        data_json = self.http.get_json_response_by_url(url, get_value=True)

        return data_json

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

    def get_email_by_id(self, email_id, mailbox_address):
        email_json = self.__read_email_by_id(email_id, mailbox_address)
        #print(email_json)
        email_df = pd.json_normalize(email_json)
        return email_json


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
            folder_id = self.MailboxFolderService.get_folder_id(mailbox_address, folder)
            folder_path = f"/mailFolders/{folder_id}" 
        else:
            folder_path = ""


        filter_suffix = self.__build_email_query_params(
            subject, sender, n_of_messages, has_attachments, greater_than_date, less_than_date
        )
        
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox_address}{folder_path}/messages?{filter_suffix}"
        

        data_json = self.http.get_json_response_by_url(url, get_value=True)
        
        if messages_json_path:
            try:
                with open(messages_json_path, "w+", encoding="utf-8") as file:
                    json.dump(data_json, file, ensure_ascii=False, indent=4)
            except Exception as e:
                raise self.HermesMSGraphError(
                    "Unable to save messages to messages.json file"
                ) from e
            
        return data_json

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

    def download_attachment(
        self, mailbox_address, email_id, attachment_id, file_name
    ):
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox_address}/messages/{email_id}/attachments/{attachment_id}"
        data_json = self.http.get_json_response_by_url(url, get_value=False)

        if data_json:
            directory = os.path.dirname(file_name)
            os.makedirs(directory, exist_ok=True)

            with open(file_name, "wb") as file:
                file.write(base64.b64decode(data_json["contentBytes"]))
            print(f"Attachment saved as {file_name}")


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