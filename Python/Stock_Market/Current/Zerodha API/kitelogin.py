import os
import sys
import json
import datetime
from kiteconnect import KiteConnect
from baselogger import logger

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
                logger.debug("Loading existing credentials...")
                creds = json.load(file)
        else:
            logger.debug("First time - going to enter the credentials...")
            creds = {
                "api_key": input("Enter your API Key: "),
                "api_secret": input("Enter your API Secret: ")
            }
            os.makedirs(os.path.dirname(self.credentials_path), exist_ok=True)
            with open(self.credentials_path, "w") as file:
                json.dump(creds, file, indent=4)
            logger.debug("Credentials saved successfully.")
        self.api_key = creds["api_key"]
        return creds

    def _get_access_token(self):
        token_file = f"Login_Credentials/{datetime.date.today()}.txt"
        if os.path.exists(token_file):
            with open(token_file, "r") as file:
                self.access_token = file.read().strip()
            logger.debug("Access token loaded from file.")
        else:
            kite = KiteConnect(api_key=self.login_credentials["api_key"])
            msg = f"Login to Zerodha account to get the access token. Open the following URL in your browser: {kite.login_url()}"
            logger.debug(msg)            
            print(msg)
            request_token = input("Enter the request token from the URL: ")
            try:
                self.access_token = kite.generate_session(
                    request_token, self.login_credentials["api_secret"])['access_token']
                logger.debug("Access token generated successfully.")
                with open(token_file, "w") as file:
                    file.write(str(self.access_token))
                logger.debug("Access token saved successfully.")
            except Exception as e:
                logger.error(f"Error generating access token: {e}")
                sys.exit(1)

    def _init_kite(self):
        self.kite = KiteConnect(api_key=self.api_key)
        self.kite.set_access_token(self.access_token)
        logger.debug("Kite Connect initialized successfully.")

    def get_api_key(self):
        return self.api_key

    def get_access_token(self):
        return self.access_token