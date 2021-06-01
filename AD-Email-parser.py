from imapclient import IMAPClient
import mailparser
from bs4 import BeautifulSoup
import re
from pyad import pyad
import logging
from pyad import adquery
import time
import requests


trial = 0
HOST = 'imap-mail.outlook.com'
PORT = 993
USERNAME = ''
PASSWORD = ''
#access_token = ''
Group_Name = ''
Group_DN = ''
url = ""


#function to get the groups that a user belongs to
def get_user_groups(func_email):
    q = adquery.ADQuery()
    q.execute_query(
        attributes=["distinguishedName", "description", "memberOf"],
        where_clause="mail = '{}'".format(func_email),
        base_dn="OU=, DC=, DC="
    )
    for row in q.get_results():
        gp = row["memberOf"]
    return gp


#function to get the user's DN from their email address
def get_user_dn(func_email):
    q = adquery.ADQuery()
    q.execute_query(
        attributes=["distinguishedName", "description"],
        where_clause="mail = '{}'".format(func_email),
        base_dn="OU=, DC=, DC="
    )
    for row in q.get_results():
        dn = row["distinguishedName"]
    return dn


def get_user_name(func_email):
    q = adquery.ADQuery()
    q.execute_query(
        attributes=["distinguishedName", "description", "displayName"],
        where_clause="mail = '{}'".format(func_email),
        base_dn="OU=, DC=, DC="
    )
    for row in q.get_results():
        dn = row["displayName"]
    return dn


def send_email(func_email, full_name):
    ploads = {"recipient": "{}".format(func_email), "fullName": "{}".format(full_name)}
    r = requests.post(url=url, json=ploads, verify=False)


#log file
logging.basicConfig(filename='script.log', level=logging.INFO)

while True:
    try:
        # Attempts to connect to the Mail Server
        with IMAPClient(HOST, use_uid=True, ssl=True, port=PORT) as server:
            try:
                #server.oauth2_login(user=USERNAME,access_token=)
                server.login(USERNAME, PASSWORD)
                trial = 0
                print("Mail Server: Login Successful")
                logging.info("Mail Server: Login Successful")
            except Exception:
                print("Mail Server: Login Failed")
                logging.info("Mail Server: Login Failed")
            while True:
                server.select_folder('Inbox')
                messages = server.search(['UNSEEN', 'FROM', '', 'SUBJECT', ''])
                #For each message that you were able to find that meets the search criteria
                for uid, message_data in server.fetch(messages, 'RFC822').items():
                    email_message = mailparser.parse_from_string(message_data[b'RFC822'].decode('utf-8'))
                    soup = BeautifulSoup(email_message.body, "html.parser")
                    text = soup.get_text()
                    # Find all email addresses in the body of the email
                    email = re.findall(r'[\w\.-]+@[\w\.-]+', text)
                    print(email)
                    email = email[0]
                    # If try block fails it means user cannot be found in AD
                    try:
                        #gets the User's DN from email
                        dn = get_user_dn(email)
                        print(dn)
                        print(get_user_groups(email))
                        # Checks if the DN is in the Users group.
                        if (get_user_groups(email)) is None or Group_DN not in get_user_groups(email):
                            adgroup = pyad.from_cn(Group_Name)
                            user = pyad.from_dn(dn)
                            logging.info("Active Directory: ADDING " + str(user) + " TO " + str(adgroup))
                            adgroup.add_members([user])
                            fullname = get_user_name(email)
                            send_email(email, fullname)
                            logging.info("Active Directory: " + str(user) + " Added to " + str(Group_Name))
                        else:
                            fullname = get_user_name(email)
                            send_email(email, fullname)
                            logging.info("User with Email : " + str(email) + " is already in group " + str(Group_DN))
                    except UnboundLocalError:
                        logging.info("Error: the user with email " + str(email) + " cannot be found")
    except Exception:
        if trial > 3:
            print("Connection: Lost connection to the server")
            logging.info("Connection: Lost connection to the server")
            exit(0)

        trial += 1
        print("Connection: Lost connection, Attempting to reconnect")
        logging.info("Connection: Lost connection, Attempting to reconnect")
        # Waits 20 seconds then attempts to reconnect
        time.sleep(60)
