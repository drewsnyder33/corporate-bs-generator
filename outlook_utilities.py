from win32com.client import Dispatch
from os import getcwd
from datetime import datetime
from os.path import join


def launch_outlook_api():

    outlook = Dispatch('Outlook.Application')
    outlook_api = outlook.GetNamespace('MAPI')

    return outlook, outlook_api

def get_inbox(outlook_api):
    return outlook_api.GetDefaultFolder(6)

def get_email(email_subject, outlook_folder):
    messages = outlook_folder.Items

    # Scan outlook folder to find if target email is present
    for message in messages:

        if message.Subject == email_subject:
            return message

    # If we've searched every single message in the folder without finding one
    # with the given subject line, it's not present, thus return None
    return None

def get_attachments(message):

    attachments = []
    for attachment in message.Attachments:
        attachments.append(attachment.Item)

    return attachments

def save_attachments(message, filepath=None, add_datestamp=False):
    '''
    Note that 'filepath' argument works when the intended file location is
    specified either with forward slashes separating folders ("/"), or with
    the entire file location string as a raw string, i.e., with an "r"
    immediately preceding the file location string, and then with back
    slashes.
    '''

    if not filepath:
        filepath = getcwd()

    if add_datestamp:
        current_day = datetime.today()
        datestamp = current_day.strftime('%Y-%m-%d')

    for attachment in message.Attachments:

        if add_datestamp:
            # Separate file extension from filename
            raw_filename = attachment.FileName
            filename_split = raw_filename.split('.')

            # Rebuild filename, without extension
            filename_stem = '.'.join(filename_split[:-1])

            file_extension = filename_split[-1]

        try:
            filename = f'{filename_stem}_{datestamp}.{file_extension}'
        except NameError:
            filename = raw_filename

        attachment.SaveAsFile(join(filepath, filename))

def format_email_recipients_from_list(recipient_list):

    return '; '.join(recipient_list)

def send_email(
    outlook_session,
    subject,
    to,
    cc='',
    bcc='',
    attachment_list=[],
    body='',
    body_html='',
    display_before_sending=False
):

    message = outlook_session.CreateItem(0x0)

    message.Subject = subject
    message.To = to
    message.CC = cc
    message.BCC = bcc

    for attachment in attachment_list:
        message.Attachments.Add(attachment)

    message.Body = body
    message.HTMLBody = body_html

    if display_before_sending:
        # Display message in Outlook, to allow reviewing before sending manually
        message.display()
    else:
        # Send message automatically, without displaying
        message.Send()
