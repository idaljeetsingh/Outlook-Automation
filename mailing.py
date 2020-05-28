"""
    module: mailing
    Author: Daljeet Singh Chhabra
    Date Created: 28-05-2020
    Last Modified: 28-05-2020
    Description:
                Script to manage the mailing.
"""

from config import VOTING_OPTIONS
import win32com.client


def reply_all(mail, msg, poll=False):
    """
        Function to send automated reply to the mail item
    :param mail: The mail item on which reply has to be sent
    :param msg: The reply message text
    :param poll: Send mail with a poll.
    :return:
    """
    reply = mail.ReplyAll()
    reply.Body = f"{msg} \n {mail.body}"
    if poll:
        reply.VotingOptions = VOTING_OPTIONS

    reply.Send()
    print('Reply sent..')
