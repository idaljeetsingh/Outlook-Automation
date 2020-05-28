"""
    module: main
    Author: Daljeet Singh Chhabra
    Date Created: 21-05-2020
    Last Modified: 21-05-2020
    Description:
                Script to automate Microsoft Outlook Desktop application.
"""
from config import SUBJECT, ERR_NO_ATTACHMENT_FOUND, ALLOWED_ATTACHMENT_TYPES
from mailing import reply_all
import win32com.client
import backend
import os

last_mail_EntryID = None


def process_mail(message):
    """
        Function to process the email.
    :param message: Email to be processed
    :return: {} of email received
    """
    mail = {
        'id': message.EntryID,
        'rec_time': str(message.CreationTime.replace(microsecond=0, tzinfo=None)),
        'subject': message.subject,
        'body': message.body
    }
    print(mail['subject'])
    if check_subject(mail['subject']):
        # Continue Automation tasks
        attachments = check_attachment_and_download(message)
        if attachments:
            print('Got true from check_attachment()')
            for attachment in attachments:
                # Continue work for each attachment
                print(attachment)

                # print('Sending automated reply to the sender...')
                # reply_all_to_mail(message, f'Received your email at: {str(datetime.now().replace(microsecond=0))} and it is processed.')
                # reply_all(message, 'text')
                # reply_all(message, 'text', poll=True)

        elif attachments is False:
            # No attachment was found
            print('No attachment was received')
            # reply_all(message, ERR_NO_ATTACHMENT_FOUND)

        mail['processed'] = 'TRUE'
    else:
        print('Subject Not matched')
    return mail


def get_mail():
    """
        Function to get last mail from Outlook Application
    :return:
    """
    global last_mail_EntryID

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    outlook.SendAndReceive(6)
    inbox = outlook.GetDefaultFolder(6)  # "6" refers to the index of a folder - in this case,

    # the inbox. You can change that number to reference
    # any other folder
    messages = inbox.Items

    message = messages.GetLast()

    if not last_mail_EntryID == message.EntryID:
        mail = process_mail(message)
        backend.insert(mail['id'], mail['rec_time'], mail['subject'], 'TRUE')
        print('*' * 10)
    last_mail_EntryID = message.EntryID


def check_subject(subject):
    """
        Function to check the subject for continuing automation
    :param subject: Subject of the email
    :return: True/False
    """
    if subject == SUBJECT:
        return True
    else:
        return False


def check_attachment_and_download(mail):
    """
        Function to look for attachment along with email and download it automation
    :param mail: Email to process
    :return: True/False
    """
    print('Checking attachments')
    attachments_in_mail = mail.Attachments
    print("Total attachments found: ", len(attachments_in_mail))
    if len(attachments_in_mail) == 0:
        return False

    attachments_list = []

    for i in range(1, len(attachments_in_mail) + 1):
        attachment = attachments_in_mail.item(i)
        print('Attachment Name: ', attachment)
        if str(attachment).split('.')[-1] not in ALLOWED_ATTACHMENT_TYPES:
            return False
        attachment.saveAsFile(os.path.join(os.path.abspath('.'), str(attachment)))
        attachments_list.append(str(attachment))
    return attachments_list


def get_email_from_entryID(entry_id):
    """
        Function to get particular email by entry id
    :param entry_id: EntryID of mail
    :return:
    """
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    email = outlook.GetItemFromID(entry_id)
    print(email.subject)


def complete_pending(last_session_id):
    """
        Complete the pending task
    :param last_session_id: Last mail that was processed
    :return:
    """
    global last_mail_EntryID
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    outlook.SendAndReceive(6)
    inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
    messages = inbox.Items

    message = messages.GetLast()
    pending = []
    while not message.EntryID == last_session_id:
        mail = process_mail(message)

        pending.append(mail)
        message = messages.GetPrevious()
    if pending:
        # Database entry of pending tasks
        for task in pending[::-1]:
            global last_mail_EntryID
            backend.insert(task['id'], task['rec_time'], task['subject'], task['processed'])
            last_mail_EntryID = task['id']
    last_mail_EntryID = last_session_id
    print('last_mail_EntryID: ', last_mail_EntryID)


def start():
    """
        Starting point of app.
    """
    # Check for old mails
    last_session = backend.get_last_mail_id()  # False means First time user
    if last_session:
        print('RESUMING LAST SESSION!!')
        print(last_session)
        complete_pending(last_session)
        while True:
            get_mail()
    else:  # First time user
        print('STARTING NEW SESSION!!')
        while True:
            get_mail()


start()
