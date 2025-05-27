import os
import sys

sys.path.append(os.path.abspath(os.path.dirname(__file__)))

from http_client import HttpClient
from email_service import EmailService
from mailbox_folder_service import MailboxFolderService
from planner_service import PlannerService
from users_service import UsersService
from exceptions import HermesMSGraphError       


class HermesMSGraph:
    """
    Class to interact with the Microsoft Graph API for sending and reading emails.

    The class contains methods to obtain an access token, send emails, read email messages, organize data into a DataFrame, and save it to a JSON file.
    """

    def __init__(self, client_id, client_secret, tenant_id):
        self.http_client = HttpClient(client_id, client_secret, tenant_id)
        self.email_service = EmailService(self.http_client)
        self.folder_service = MailboxFolderService(self.http_client)
        self.planner_service = PlannerService(self.http_client)
        self.users_service = UsersService(self.http_client)
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.http = HttpClient(client_id, client_secret, tenant_id)

        # EmailService methods
    def send_email(self, sender_mail, subject, body, to_address, cc_address=None):
        return self.email_service.send_email(sender_mail, subject, body, to_address, cc_address)

    def get_emails(self, mailbox_address, **kwargs):
        return self.email_service.get_emails(mailbox_address, **kwargs)

    def move_email_to_folder(self, email_id, mailbox_address, folder_name=None, folder_id=None):
        return self.email_service.move_email_to_folder(email_id, mailbox_address, folder_name, folder_id)

    def forward_email_by_id(self, email_id, mailbox_address, to_address, comment=None):
        return self.email_service.foward_email_by_id(email_id, mailbox_address, to_address, comment)
    
    def list_email_attachments(self, email_id, mailbox_address):
        return self.email_service.list_email_attachments(email_id, mailbox_address)

    # MailboxFolderService methods
    def list_mailbox_folders(self, mailbox_address):
        return self.folder_service.list_mailbox_folders(mailbox_address)

    def get_mailbox_folders(self, mailbox_address):
        return self.folder_service.get_mailbox_folders(mailbox_address)

    def validate_folder_id(self, mailbox_address, folder_id):
        return self.folder_service.validate_folder_id(mailbox_address, folder_id)

    def get_folder_id(self, mailbox_address, folder_name):
        return self.folder_service.get_folder_id(mailbox_address, folder_name)

    # PlannerService methods
    def list_plans_by_group_id(self, group_id, data="all"):
        return self.planner_service.list_plans_by_group_id(group_id, data)

    def list_visible_plans_by_user_id(self, user_id):
        return self.planner_service.list_visible_plans_by_user_id(user_id)

    def list_tasks_by_user_id(self, user_id):
        return self.planner_service.list_tasks_by_user_id(user_id)

    def list_tasks_by_plan_id(self, planner_id):
        return self.planner_service.list_tasks_by_plan_id(planner_id)
    
    # UsersService methods
    def get_user_id_by_email(self, email_address):
        return self.users_service.get_user_id_by_email(email_address)
    
    def get_all_users(self, data='all'):
        return self.users_service.get_all_users(data)
    
    def search_from_mailboxes(self, query):
        return self.users_service.search_from_mailboxes(query)
    
    def get_tenant_licenses(self):
        return self.users_service.get_tenant_licenses()
    
    def add_user_to_shared_mailbox(self, user_address, shared_mailbox_address):
        return self.users_service.add_user_to_shared_mailbox(user_address, shared_mailbox_address)
    
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

    def __json_to_dataframe(self, json_file_path):
        with open(json_file_path, "r") as file:
            json_data = json.load(file)
        df_emails = pd.json_normalize(json_data)
        return df_emails



    # Usage example:
    # api = MSGraphAPI()
    # df_emails = api.get_df_emails("anakin.skywalker@github.com")      

    
    def list_msgraph_permissions(self):
        self.http.list_msgraph_permisions()
