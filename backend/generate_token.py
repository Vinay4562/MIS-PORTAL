import os
import json
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# Scopes required for sending email
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
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
            if not os.path.exists('credentials.json'):
                print("Error: 'credentials.json' not found.")
                print("1. Go to Google Cloud Console -> APIs & Services -> Credentials")
                print("2. Create OAuth Client ID (Desktop App)")
                print("3. Download JSON and rename to 'credentials.json' in this folder")
                return

            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
            
    print("\nSuccessfully authenticated!")
    print("Here are your credentials for Railway (add these to Railway Variables):")
    
    # Parse the token.json to get the values
    with open('token.json', 'r') as f:
        data = json.load(f)
        
    print("-" * 50)
    print(f"GOOGLE_CLIENT_ID:     {data.get('client_id')}")
    print(f"GOOGLE_CLIENT_SECRET: {data.get('client_secret')}")
    print(f"GOOGLE_REFRESH_TOKEN: {data.get('refresh_token')}")
    print("-" * 50)
    print("\nNote: You also need to set SMTP_EMAIL to your gmail address.")

if __name__ == '__main__':
    main()