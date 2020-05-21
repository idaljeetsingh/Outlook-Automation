"""
    module: backend
    Author: Daljeet Singh Chhabra
    Date Created: 21-05-2020
    Last Modified: 21-05-2020
    Description:
                Script to manage the database operations.
"""
import sqlite3


def configureDB():
    """
        Base database configuration
    :return:
    """
    conn = sqlite3.connect("mails.db")
    cur = conn.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS mails (id TEXT, rec_on TEXT, subj TEXT, processed TEXT)")
    conn.commit()
    conn.close()


def insert(id, rec_on, subj, processed):
    """
        Inserts the last mail record in DB
    :param id: Mail's EntryID
    :param rec_on: Receive time of mail
    :param subj: Subject of Mail
    :param processed: Processed type of mail
    :return:
    """
    conn = sqlite3.connect("mails.db")
    cur = conn.cursor()
    cur.execute(f'INSERT INTO mails VALUES("{id}","{rec_on}","{subj}","{processed}")')
    conn.commit()
    conn.close()


def get_last_mail_id():
    """
        Fetch last email's EntryID
    :return: EntryID
    """
    conn = sqlite3.connect("mails.db")
    cur = conn.cursor()
    cur.execute("SELECT id FROM mails ORDER BY id DESC LIMIT 1;")
    row = cur.fetchone()
    conn.close()
    if row:
        return row[0]
    else:
        return False


i = '00000000A2A07039A98390418EC5D07FE4E9359E07009601036F7B945B489698B3AA62F8576900000000010C00009601036F7B945B489698B3AA62F857690001F508D5280000'

configureDB()
# insert(i, "21-05-2020", "Dummy", "TRUE")
# print(get_last_mail_id())
