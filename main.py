import os

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from datetime import date,timedelta
import smtplib
from googleapiclient.discovery import build
from email.mime.text import MIMEText
from google.auth.transport.requests import Request
from googleapiclient.errors import HttpError

#Deleted For Company Information
sender_email = ""
recipients = ['', '']
test_mail = ""
#Deleted For Company Information
booking_list_id = ""
#Deleted For Company Infromation
sheet_id = ""
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
range = ""
#Deleted For Company Information
booking_list_range = ""


def API_Connection():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=booking_list_id,range=booking_list_range).execute()
        data = result.get('values', [])
        if not data:
            print('No data found.')
            return

        return data

    except HttpError as err:
        print(err)
def sendMail(bookingNumber,date):
    message = str(bookingNumber) + " için Hatırlatma mailidir. Cut Off Tarihi : " + str(date)
    msg = MIMEText(message)
    msg['Subject'] = str(bookingNumber) + ' İÇİN HATIRLATMA MAİLİ'
    msg['From'] = sender_email
    msg['To'] = ", ".join(recipients)
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    #Deleted For Company Information
    smtp_username = ''
    smtp_password = ''

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.sendmail(sender_email, recipients, msg.as_string())


if __name__ == '__main__':
    values = API_Connection()
    today = date.today()
    tomorrow = today + timedelta(days=1)
    for row in values:
        if(row != 0):
            if ((len(row)>9) and(row[9] == today.strftime("%d.%m"))):
                if((len(row)>14 and row[14] == '') or (len(row)>11 and not row[11]=='OK')):
                        sendMail(row[4],"BUGÜN")
                        print(row)

