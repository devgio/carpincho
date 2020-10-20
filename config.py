# GENERAL
download_folder = r'C:\example_folder'

# MAILBOX SETTINGS
from src.email_connectors.O365 import Mailbox as mailbox_class # Set the type of mailbox connector to use, so far the only available is O365
mailbox_connection_parameters = ('credentials', 'credentials') # If the Azure credentials need to be updated, refer to https://pypi.org/project/O365/#authentication
shared_mailbox = False # If not using a shared mailbox, set this to False
inbox = ['Inbox']
error_folder = ['Script_Error']
unrecognized_folder = ['Unrecognized']
