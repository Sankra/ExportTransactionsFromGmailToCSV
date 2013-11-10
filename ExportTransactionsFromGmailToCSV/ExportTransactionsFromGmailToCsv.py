#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# This script parses emails in Gmail to find financial transactions
# and exports them to a .csv-file.
# 
# Instructions
# - easy_install beautifulsoup4
# - Add Gmail username, password and a financial account name
# - (modify parsers to suit your needs)
# - Run the script!
#
# Author: Runar Ovesen Hjerpbakk - http://hjerpbakk.com

import imaplib
import email
import csv
import os
import locale
from bs4 import BeautifulSoup

user = ''
password = ''
account = ''
csv_path = os.path.expanduser('transactions.csv')
csv_text_encoding = 'iso-8859-1'
category_column = 3
amount_column = 4
no = 'no_NO'
en = 'en_US'

def archive_email(imap, uid):
    imap.uid('STORE', uid, '+FLAGS', r'(\Seen)')
    imap.uid('COPY', uid, '[Gmail]/All Mail')
    imap.uid('STORE', uid, '+FLAGS', '\\Deleted')
    imap.expunge()

def create_negative_amount(amount):
    # Remove trailing 'kr' and convert to float
    global current_locale   
    if ('.' in amount):
        locale.setlocale(locale.LC_ALL, en)
    else:
        locale.setlocale(locale.LC_ALL, no)
    value = locale.atof(amount[:-2])
    return '%.2f' % -value

def get_itunes_values(column, text):
    if column == category_column:
        if text == 'App' or text == 'iOS App' or text == 'In App Purchase':
            return 'iTunes > Apps'
        if text == 'Song' or text == 'Playlist':
            return 'iTunes > Music'
        if text == 'Book':
            return 'iTunes > Books'
        if text == 'Subscription Renewal' or text == 'Tone' or text == 'Ringtone' or text == 'Init. Subscription':
            return 'iTunes > Other'
        if text == 'Film (HD)' or text == 'Film' or text == 'Video':
            return 'iTunes > Movies'
    if column == amount_column:
        return create_negative_amount(text)
    return text

def parse_itunes_transactions(mail_body):
    soup = BeautifulSoup(mail_body)
    receipt = soup.findChildren('table')[2]
    date = list(receipt.findChildren('td')[2].stripped_strings)[3]
    items = []
    number_of_categories = 5
    i = 0
    transaction_table = soup.findChildren('table')[4]
    transactions = transaction_table.findChildren(['tr'])
    for transaction in transactions:
        cells = transaction.findChildren('td')
        row_values = [account]
        for cell in cells:
            values = list(cell.stripped_strings)
            if len(values) > 0:
                text = values[0].encode(csv_text_encoding)
                if (i > 0):
                    row_values.append(get_itunes_values(len(row_values), text))
                else:
                    row_values.append(text)
        if len(row_values) == number_of_categories:
            if i > 0:
                row_values.append(date)
                items.append(row_values)
            i += 1
    return items

def get_decoded_email_body(message_body):
    msg = email.message_from_string(message_body)
    text = ""
    if msg.is_multipart():
        html = None
        for part in msg.get_payload():
            if part.get_content_charset() is None:
                # We cannot know the character set, so return decoded "something"
                text = part.get_payload(decode=True)
                continue
            charset = part.get_content_charset()
            if part.get_content_type() == 'text/plain':
                text = unicode(part.get_payload(decode=True), str(charset), "ignore").encode('utf8', 'replace')
            if part.get_content_type() == 'text/html':
                html = unicode(part.get_payload(decode=True), str(charset), "ignore").encode('utf8', 'replace')
        if html is not None:
            return html.strip()
        else:
            return text.strip()
    else:
        text = unicode(msg.get_payload(decode=True), msg.get_content_charset(), 'ignore').encode('utf8', 'replace')
        return text.strip()

def get_transactions_and_archive(search_filter, parse_transactions):
    transactions = []
    imap = imaplib.IMAP4_SSL("imap.gmail.com", 993)
    imap.login(user, password)
    imap.select('Inbox', readonly = False)
    typ, data = imap.uid('search', None, search_filter)
    try:
        uids = data[0].split()
        for uid in uids:
            typ, msg_data = imap.uid('fetch', uid, '(RFC822)')
            for response_part in msg_data:
                if isinstance(response_part, tuple):              
                    body = get_decoded_email_body(response_part[1])
                    transactions.extend(parse_itunes_transactions(body))
        for uid in uids:
            archive_email(imap, uid)
    finally:
        try:
            imap.close()
        except:
            pass
        imap.logout()
    return transactions

def get_csv(csv_file):
    return csv.writer(csv_file, quotechar='"', quoting=csv.QUOTE_ALL)

def create_csv_if_needed():
    if (os.path.exists(csv_path)):
        return
    csv_file  = open(csv_path, 'wb')
    csv_writer = get_csv(csv_file)
    csv_writer.writerow(['Account','Description', 'Payee', 'Category', 'Amount', 'Date'])
    csv_file.close()

def write_transactions_to_csv(transactions):
    csv_file  = open(csv_path, 'ab')
    csv_writer = get_csv(csv_file)
    for transaction in transactions:
        csv_writer.writerow(transaction)
    csv_file.close()

if __name__ == "__main__":
    create_csv_if_needed()
    transactions = get_transactions_and_archive('(FROM do_not_reply@itunes.com) (SUBJECT Receipt)', parse_itunes_transactions)
    write_transactions_to_csv(transactions)
