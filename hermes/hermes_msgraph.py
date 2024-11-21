#!/usr/bin/env python
# https://graph.microsoft.com/v1.0/me/messages?$filter=subject eq '{subject}' and sender/emailAddress/address eq '{sender email address}' and sentDateTime ge 2023-05-17T07:28:08Z
# coding: utf-8
import json
import requests
import yaml
import os
import base64
import pandas as pd

class MSGraphAPI:
    """
    Class to interact with the Microsoft Graph API for sending and reading emails.

    The class contains methods to obtain an access token, send emails, read email messages, organize data into a DataFrame, and save it to a JSON file.
    """

    def __init__(self, client_id, client_secret, tenant_id):
        """
        Initializes the class with the credentials passed as parameters.

        Parameters:
        ----------
        client_id : str
            Client ID registered in Azure AD.
        client_secret : str
            Client secret registered in Azure AD.
        tenant_id : str
            Directory (tenant) ID in Azure AD.
        """
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id

    def __get_access_token(self):
        """
        Private method to obtain the access token for the Microsoft Graph API.

        Returns:
        -------
        access_token : str
            Access token to make API calls to Microsoft Graph.

        Raises:
        ------
        RuntimeError:
            If the request fails when trying to obtain the access token.
        """
        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        payload = {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": "https://graph.microsoft.com/.default"
        }
        try:
            response = requests.post(url, data=payload)
        except Exception as e:
            raise RuntimeError("Unable to make post request to get access token") from e
        
        response_data = response.json()
        access_token = response_data['access_token']
        return access_token

    def send_email(self, sender_mail, subject, body, to_address, cc_address=None):
        """
        Sends an email using the Microsoft Graph API.

        Parameters:
        ----------
        sender_mail : str
            Sender's email address.
        subject : str
            Email subject.
        body : str
            Email body.
        to_address : str
            Recipient's email address.
        cc_address : str, optional
            CC email address, default is None.

        Returns:
        -------
        None
        """
        access_token = self.__get_access_token()
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }

        url = f"https://graph.microsoft.com/v1.0/users/{sender_mail}/sendMail"

        payload = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "Text",
                    "content": body
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": to_address
                        }
                    }
                ],
                "ccRecipients": [
                    {
                        "emailAddress": {
                            "address": cc_address
                        }
                    }
                ] if cc_address else []
            },
            "saveToSentItems": "true"
        }

        response = requests.post(url, headers=headers, data=json.dumps(payload))

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

    def __filter_subject(self, df, subject_filter):
        """
        Filters the DataFrame for email messages matching the subject filter.

        Parameters:
        ----------
        df : pandas.DataFrame
            DataFrame of email messages.
        subject_filter : str
            Subject filter that may use asterisks for partial matching.

        Returns:
        -------
        df : pandas.DataFrame
            DataFrame filtered by subject.
        """
        if subject_filter:
            # Case: starts and ends with asterisk -> "contains"
            if subject_filter.startswith("*") and subject_filter.endswith("*"):
                subject_filter = subject_filter.strip("*")
                df = df[df["subject"].str.contains(subject_filter, case=False, regex=False)]
            
            # Case: starts with asterisk -> "ends with"
            elif subject_filter.startswith("*") and not subject_filter.endswith("*"):
                subject_filter = subject_filter.lstrip("*")
                df = df[df["subject"].str.endswith(subject_filter, na=False)]
            
            # Case: ends with asterisk -> "starts with"
            elif subject_filter.endswith("*") and not subject_filter.startswith("*"):
                subject_filter = subject_filter.rstrip("*")
                df = df[df["subject"].str.startswith(subject_filter, na=False)]
            
            # Case: exact match
            else:
                df = df[df["subject"] == subject_filter]

        return df

    def list_email_attachments(self, email_address, message_id):
        access_token = self.__get_access_token()
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        
        url = f"https://graph.microsoft.com/v1.0/users/{email_address}/messages/{message_id}/attachments"
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            return response.json().get('value', [])
        else:
            print(f"Error: {response.status_code}")
            return None
    
    def download_attachment(self, email_address, message_id, attachment_id, file_name):
        access_token = self.__get_access_token()  # Função para obter o token de acesso
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        
        # URL para acessar o anexo
        url = f"https://graph.microsoft.com/v1.0/users/{email_address}/messages/{message_id}/attachments/{attachment_id}"
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            attachment_data = response.json()
            
            # Criar diretório se não existir
            directory = os.path.dirname(file_name)
            os.makedirs(directory, exist_ok=True)
            
            # Salvar o conteúdo do anexo no arquivo
            with open(file_name, "wb") as file:
                file.write(base64.b64decode(attachment_data['contentBytes']))
                #file.write(bytes(attachment_data['contentBytes'], 'utf-8'))
            print(f"Attachment saved as {file_name}")
        else:
            print(f"Error: {response.status_code}")


    def get_email_folders(self, email_address):
        """
        Retrieves the email folders for the specified user.

        Parameters:
        ----------
        email_address : str
            User's email address.

        Returns:
        -------
        list
            A list containing the email folders.

        Raises:
        ------
        RuntimeError:
            If an error occurs when accessing the email folders.
        """
        access_token = self.__get_access_token()
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        url = f"https://graph.microsoft.com/v1.0/users/{email_address}/mailfolders/delta?$select=displayname"
        
        #https://stackoverflow.com/questions/42901755/microsoft-graph-outlook-mail-list-all-mail-folders-not-just-the-top-level-o
        #url = f"https://graph.microsoft.com/v1.0/users/{email_address}/messages?$top={n_of_messages}"

        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            mail_folders = response.json().get('value', [])
            return mail_folders
        else:
            print(f"Error reading emails: {response.status_code}")

    def list_email_folders(self, email_address):
        """
        Lists email folders and returns a DataFrame.

        Parameters:
        ----------
        email_address : str
            User's email address.

        Returns:
        -------
        pandas.DataFrame
            DataFrame containing the email folders.
        """
        mail_folders = self.get_email_folders(email_address)
        df_mail_folders = pd.DataFrame(mail_folders)
        if not df_mail_folders.empty:
            return df_mail_folders

    def __read_email(self, email_address, subject, folder, sender, n_of_messages, has_attachments, subject_filter, messages_json_path):
        """
        Private method to read emails from a specific address.

        Parameters:
        ----------
        email_address : str
            Email address to read messages from.
        subject : str
            Subject to filter.
        folder : str
            Folder name to filter messages.
        sender : str
            Sender's address.
        n_of_messages : int
            Number of messages to retrieve.
        subject_filter : str
            Subject filter to apply to messages.
        messages_json_path : str
            Path to the JSON file where messages will be saved.

        Returns:
        -------
        None
        """
        email_filter = []
        email_filter_url = []

        if folder:
            df_folders = self.list_email_folders(email_address)
            folder_id = df_folders.loc[df_folders['displayName'] == folder, 'id'].iloc[0]
            folder= f"/mailFolders/{folder_id}"
        
        if sender:
            email_filter.append(f"sender/emailAddress/address eq '{sender}'")
        
        if subject:
            email_filter.append(self.__ulr_filter_subject(subject))
            #email_filter.append(f"subject eq '{subject}'")
        if has_attachments:
            email_filter.append(f"hasAttachments eq {str(has_attachments).lower()}")

        if email_filter:
            email_filter_joined = " and ".join(email_filter)
            email_filter_url.append(f"$filter={email_filter_joined}")

        if n_of_messages:
            email_filter_url.append(f"$top={n_of_messages}")
                
        access_token = self.__get_access_token()
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        if email_filter_url:
            email_filter_url = "&".join(email_filter_url)
        else:
            email_filter_url = ''

        url = f"https://graph.microsoft.com/v1.0/users/{email_address}{folder}/messages?{email_filter_url}"
        #print(url)
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            messages = response.json().get('value', [])
            try:
                with open(messages_json_path, 'w+', encoding='utf-8') as file:
                    json.dump(messages, file, ensure_ascii=False, indent=4)
            except Exception as e:
                raise RuntimeError("Unable to save messages to messages.json file") from e
        else:
            print(f"Error reading emails: {response.status_code}")

    def __json_to_dataframe(self, json_file_path):
        """
        Private method to convert a JSON file of email messages into a DataFrame.

        Parameters:
        ----------
        json_file_path : str
            Path to the JSON file with email messages.

        Returns:
        -------
        df_emails : pandas.DataFrame
            DataFrame containing email messages.
        """
        with open(json_file_path, 'r') as file:
            json_data = json.load(file)
        df_emails = pd.json_normalize(json_data)
        return df_emails

    def __organize_df_emails(self, df_emails):
        """
        Private method to organize the email messages DataFrame with relevant columns.

        Parameters:
        ----------
        df_emails : pandas.DataFrame
            DataFrame with email messages.

        Returns:
        -------
        df_emails : pandas.DataFrame
            Organized DataFrame with specific columns.
        """
        columns = ['subject', 'isRead', 'sentDateTime', 'receivedDateTime', 'sender.emailAddress.name', 
                   'sender.emailAddress.address', 'from.emailAddress.name', 'from.emailAddress.address', 
                   'bodyPreview', 'body.contentType', 'body.content', 'id']
        df_emails = df_emails[columns]
        return df_emails

    def get_df_emails(self, email_address, subject='', folder='', sender='', n_of_messages=10, has_attachments = '', subject_filter="", messages_json_path='messages.json'):
        """
        Public method to retrieve email messages and organize them into a DataFrame with only essential columns, excluding all email attributes.

        Parameters:
        ----------
        email_address : str
            Email address from which to retrieve messages.
        subject : str, optional
            Subject to filter by, default is an empty string.
        folder : str, optional
            Folder to retrieve messages from, default is an empty string.
        sender : str, optional
            Sender's email address to filter by, default is an empty string.
        n_of_messages : int, optional
            Number of messages to retrieve, default is 10.
        subject_filter : str, optional
            Filter to apply to email subjects, default is an empty string.
        messages_json_path : str, optional
            Path to the JSON file where emails will be saved, default is 'messages.json'.

        Returns:
        -------
        df_emails : pandas.DataFrame
            Organized DataFrame with email messages.
        """
        self.__read_email(email_address, subject, folder, sender, n_of_messages, has_attachments, subject_filter, messages_json_path)
        df_raw_emails = self.__json_to_dataframe(messages_json_path)
        if not df_raw_emails.empty:
            df_emails = self.__organize_df_emails(df_raw_emails)
        else:
            #print("No emails found using the parameters passed.")
            pass
        if subject_filter != "":
            df_emails = self.__filter_subject(df_emails, subject_filter)
        return df_emails
    # Usage example:
    # api = MSGraphAPI()
    # df_emails = api.get_df_emails("anakin.skywalker@github.com")

    def get_raw_df_emails(self, email_address, subject='', folder='', sender='', n_of_messages=10, has_attachments = '', subject_filter="", messages_json_path='messages.json'):
        """
        Public method to retrieve email messages and organize them into a DataFrame with all received attributes.

        Parameters:
        ----------
        email_address : str
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
        self.__read_email(email_address, subject, folder, sender, n_of_messages, has_attachments, subject_filter, messages_json_path)
        df_raw_emails = self.__json_to_dataframe(messages_json_path)
        if subject_filter != "":
            df_raw_emails = self.__filter_subject(df_raw_emails, subject_filter)
        return df_raw_emails
