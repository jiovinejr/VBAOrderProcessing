import pandas as pd
import os.path
import base64
import ctypes
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email.mime.text import MIMEText

# Big thanks to google documentation and examples

# 
# RE DOWNLOAD CREDENTIALS.JSON
# 

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/gmail.compose"]

def create_message(to, message_text):
    """Create a message for an email."""
    message = MIMEText(message_text)
    message['to'] = to
    message['subject'] = message_text
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return {'raw': raw_message}


def main():
  """Shows basic usage of the Gmail API.
  Lists the user's Gmail labels.
  """
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    # Call the Gmail API
    service = build("gmail", "v1", credentials=creds)
    # draft = service.users().labels().list(userId="me").execute()
    # labels = results.get("labels", [])

    df = pd.read_excel('VBAtoGoogleTest.xlsx', header=None, usecols='A')

    sender = 'jiovinejr@gmail.com'
    to = 'anna@iovine.com, k.quigley@delshipusa.com, vinnie@iovine.com'
    # subject = 'Subject of the Email'
    # message_text = 'This is the body of the email.'
    drafts = []
    for index, row in df.iterrows():
        ship_name = row[0]

        # Create the email message
        message_body = create_message(to, ship_name)

        message = {'message': message_body}
        draft = service.users().drafts().create(userId="me", body=message).execute()
        print(f'Draft id: {draft["id"]}\nDraft message: {draft["message"]}')
        drafts.append(draft)
    ctypes.windll.user32.MessageBoxW(0, "Email setup complete.", "Email", 1)
    return drafts
  
  except HttpError as error:
    # TODO(developer) - Handle errors from gmail API.
    print(f"An error occurred: {error}")


if __name__ == "__main__":
  main()
