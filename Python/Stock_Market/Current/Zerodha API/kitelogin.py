import os
import sys
import json
import datetime
from kiteconnect import KiteConnect

class KiteLogin:
    def __init__(self, credentials_path="Login_Credentials/credentials.json"):
        self.credentials_path = credentials_path
        self.access_token = None
        self.api_key = None
        self.kite = None
        self.login_credentials = self._load_credentials()
        self._get_access_token()
        self._init_kite()

    def _load_credentials(self):
        if os.path.exists(self.credentials_path):
            with open(self.credentials_path, "r") as file:
                creds = json.load(file)
        else:
            print("Enter your credentials...")
            creds = {
                "api_key": input("Enter your API Key: "),
                "api_secret": input("Enter your API Secret: ")
            }
            os.makedirs(os.path.dirname(self.credentials_path), exist_ok=True)
            with open(self.credentials_path, "w") as file:
                json.dump(creds, file, indent=4)
            print("Credentials saved successfully.")
        self.api_key = creds["api_key"]
        return creds

    def _get_access_token(self):
        token_file = f"Login_Credentials/{datetime.date.today()}.txt"
        if os.path.exists(token_file):
            with open(token_file, "r") as file:
                self.access_token = file.read().strip()
        else:
            kite = KiteConnect(api_key=self.login_credentials["api_key"])
            print("Please login to your Zerodha account to get the access token.")
            print("Open the following URL in your browser:")
            print(kite.login_url())
            request_token = input("Enter the request token from the URL: ")
            try:
                self.access_token = kite.generate_session(
                    request_token, self.login_credentials["api_secret"])['access_token']
                print("Access token generated successfully.")
                with open(token_file, "w") as file:
                    file.write(str(self.access_token))
                print("Access token saved successfully.")
            except Exception as e:
                print(f"Error generating access token: {e}")
                sys.exit(1)

    def _init_kite(self):
        self.kite = KiteConnect(api_key=self.api_key)
        self.kite.set_access_token(self.access_token)
        print("Kite Connect initialized successfully.")

    def get_api_key(self):
        return self.api_key

    def get_access_token(self):
        return self.access_token

# Example usage:
# kite_login = KiteLogin()
# api_key = kite_login.get_api_key()
# access_token = kite_login.get_access_token()