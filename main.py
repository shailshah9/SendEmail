#!/usr/bin/python3

from __future__ import print_function
import httplib2
import os
import constants as c
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'secret.json'
APPLICATION_NAME = 'Google Sheets API Python Quickstart'


'''def send_email(to, name_of_recruiter, company_name):
    gmail_user = c.gmail_user
    gmail_password = c.gmail_password

    sent_from = gmail_user

    subject = c.subject % name_of_recruiter

    email_text = c.body%company_name

    print(email_text)
    #foo.encode('ascii', 'ignore')

    message = 'Subject: {}\n\n{}'.format(subject, email_text)

    try:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.ehlo()
        server.login(gmail_user, gmail_password)
        server.sendmail(sent_from, to, message)
        server.close()

        print('Email sent!')

        update_excel()

    except Exception as e:
        print(e)
        print('Something went wrong...')'''

def send_email(to,name,company_name):
    try:

        fromaddr = c.gmail_user
        toaddr = to

        # instance of MIMEMultipart
        msg = MIMEMultipart()

        # storing the senders email address
        msg['From'] = fromaddr

        # storing the receivers email address
        msg['To'] = toaddr

        # storing the subject
        msg['Subject'] = c.subject % company_name

        # string to store the body of the mail
        body = c.body % (name,company_name)

        # attach the body with the msg instance
        msg.attach(MIMEText(body, 'plain'))

        # open the file to be sent
        filename = c.resume
        attachment = open(filename, "rb")

        # instance of MIMEBase and named as p
        p = MIMEBase('application', 'octet-stream')

        # To change the payload into encoded form
        p.set_payload((attachment).read())

        # encode into base64
        encoders.encode_base64(p)

        p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

        # attach the instance 'p' to instance 'msg'
        msg.attach(p)

        # creates SMTP session
        s = smtplib.SMTP('smtp.gmail.com', 587)

        # start TLS for security
        s.starttls()

        # Authentication
        s.login(fromaddr, c.gmail_password)

        # Converts the Multipart msg into a string
        text = msg.as_string()

        # sending the mail
        s.sendmail(fromaddr, toaddr, text)

        # terminating the session
        s.quit()
        print("email sent")

    except Exception as e:
        print(e)

'''def append(spreadsheetId,service,com_name,role,app_date,email_id):
        list = [[com_name], [role], ["Applied"],[app_date], [email_id]]
        resource = {
          "majorDimension": "COLUMNS",
          "values": list
        }
        #spreadsheetId = ""
        range = "Shail!A:A";
        service.spreadsheets().values().append(
          spreadsheetId=spreadsheetId,
          range=range,
          body=resource,
          valueInputOption="USER_ENTERED"
        ).execute()'''



def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def get_last_row():
    f = open("last_row", "r")
    return int(f.readline())


def update_last_row(row_count):
    f = open("last_row","w+")
    f.write(str(row_count))
    f.close()
def read_google_sheets():
    """Shows basic usage of the Sheets API.

    Creates a Sheets API service object and prints the names and majors of
    students in a sample spreadsheet:
    https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)
    spreadsheetId=c.spreadsheetId
    #append(spreadsheetId,service,com_name,role,date_app,email)

    rangeName = 'contact!A2:C'
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        print('Name, Company, email')
        int_count = get_last_row()
        #print(int_count)
        curr_count=1
        for row in values:
            if curr_count >= int_count :
                print('%s %s, %s, %s' % (str(curr_count),row[0], row[1],row[2]))
                send_email(row[2],row[0],row[1])
            curr_count += 1

        update_last_row(curr_count-1)
        #print(get_last_row())
        print("Sent emails to %s recruiters" %(str(curr_count-int_count-1)))


if __name__ == '__main__':
    read_google_sheets()



