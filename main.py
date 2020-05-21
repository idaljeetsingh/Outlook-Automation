"""
    module: main
    Author: Daljeet Singh Chhabra
    Date Created: 21-05-2020
    Last Modified: 21-05-2020
    Description:
                Script to automate Microsoft Outlook Desktop application.
"""
import win32com.client
import backend

last_mail_EntryID = None


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
    mail = {
        'id': message.EntryID,
        'rec_time': str(message.CreationTime.replace(microsecond=0, tzinfo=None)),
        'subject': message.subject,
        'body': message.body
    }
    if not last_mail_EntryID == mail['id']:
        backend.insert(mail['id'], mail['rec_time'], mail['subject'], 'TRUE')
        print(mail)
        print('*' * 10)
    last_mail_EntryID = mail['id']


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
        mail = {
            'id': message.EntryID,
            'rec_time': str(message.CreationTime.replace(microsecond=0, tzinfo=None)),
            'subject': message.subject,
            'body': message.body
        }
        print(mail['subject'])

        processed = True
        # if processed:
        if processed:
            mail['processed'] = 'TRUE'
        else:
            mail['processed'] = 'FALSE'

        pending.append(mail)
        message = messages.GetPrevious()
    print(pending)
    print(pending[::-1])
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
