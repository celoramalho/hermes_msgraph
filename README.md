# Microsoft Graph Python API class
Project Description

MSGraph Email Manager is a Python project designed to facilitate email management and interaction with Microsoft Graph API. It provides an easy-to-use interface to automate various email-related tasks, such as sending emails, reading inbox messages, organizing data into structured formats like DataFrames, and extracting specific email data as needed.

Built with developers in mind, this package is particularly useful for businesses and individuals who want to automate email workflows, extract email data for reporting, or conduct analysis on email messages without needing to directly handle complex API calls. With comprehensive functions and documentation, MSGraph Email Manager provides the tools necessary to streamline your email operations via the Microsoft Graph API.
Key Features

    Send Emails: Automate the process of sending emails directly from a Microsoft account.
    Read Emails: Retrieve and save emails in JSON format for easy access and data analysis.
    Data Extraction: Extract and organize email data into DataFrames for quick access to essential email fields (subject, sender, date, etc.).
    Custom Filtering: Organize emails based on filters like subject or sender.
    Modular Design: Structured as a Python class with public and private methods for better control and customization.
    Easy Integration: The package can be installed and integrated in any Python environment, allowing for further automation and functionality.

Installation

    Clone the repository from GitHub:

    bash

git clone https://github.com/your-username/MSGraph-Email-Manager.git

Install the package:

bash

    cd MSGraph-Email-Manager
    pip install -e .

Example Usage

python

from msgraph_email_manager import MSGraphAPI

# Initialize the MSGraphAPI class
graph_api = MSGraphAPI(client_id="YOUR_CLIENT_ID", client_secret="YOUR_CLIENT_SECRET", tenant_id="YOUR_TENANT_ID")

# Send an email
graph_api.send_email(
    sender_mail="your-email@domain.com",
    subject="Test Email",
    body="This is a test email sent using MSGraphAPI.",
    to_address="recipient@domain.com"
)

# Retrieve recent emails
df_emails = graph_api.get_df_emails(email_address="your-email@domain.com", n_of_massages=10)
print(df_emails)

Requirements

    Python 3.6+
    requests library for handling HTTP requests
    pandas for data manipulation
    yaml for reading configuration files (optional)

License

This project is licensed under the MIT License.
