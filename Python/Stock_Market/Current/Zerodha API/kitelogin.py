import os
import sys
import json
import datetime
from kiteconnect import KiteConnect
from baselogger import logger

class KiteLogin:
    def __init__(self, credentials_path="Login_Credentials/credentials.json"):
        self.credentials_path = credentials_path
        self.api_key = None
        self.api_secret = None
        self.access_token = None
        self.kite = None
        #self.init_kite()

    def load_credentials(self):
        creds = {}
        if os.path.exists(self.credentials_path):
            with open(self.credentials_path, "r") as file:
                logger.debug("Loading existing credentials...")
                creds = json.load(file)
            self.api_key = creds["api_key"]
            self.api_secret = creds["api_secret"]
            logger.debug("Credentials loaded successfully.")
        else:
            logger.debug("No existing credentials found.")
        #return creds

    def create_credentials(self, mode="non_gui", api_key=None, api_secret=None):
        logger.debug("First time - going to enter the credentials...")
        if mode == "non_gui":
            creds = {
                "api_key": input("Enter your API Key: "),
                "api_secret": input("Enter your API Secret: ")
            }
        elif mode == "gui":
            if not api_key or not api_secret:
                logger.error("API Key and Secret must be provided in GUI mode.")
                sys.exit(1)
            creds = {
                "api_key": api_key,
                "api_secret": api_secret
            }
        else:
            logger.error("Invalid mode. Choose 'gui' or 'non_gui'.")
            sys.exit(1)
        os.makedirs(os.path.dirname(self.credentials_path), exist_ok=True)
        with open(self.credentials_path, "w") as file:
            json.dump(creds, file, indent=4)
        logger.debug("Credentials saved successfully.")
        self.api_key = creds["api_key"]
        self.api_secret = creds["api_secret"]
        #return creds

    def load_access_token(self):
        token_file = f"Login_Credentials/{datetime.date.today()}.txt"
        if os.path.exists(token_file):
            with open(token_file, "r") as file:
                self.access_token = file.read().strip()
            logger.debug("Access token loaded from file.")
        #return self.access_token

    def create_access_token(self, mode="non_gui", request_token=None):
        token_file = f"Login_Credentials/{datetime.date.today()}.txt"  
        kite = KiteConnect(api_key=self.api_key)     
        if mode == "non_gui":            
            msg = f"Login to Zerodha account to get the access token. Open the following URL in your browser: {kite.login_url()}"
            logger.debug(msg)            
            print(msg)
            request_token = input("Enter the request token from the URL: ")
        elif mode == "gui":
            if not request_token:
                logger.error("Request Token must be provided in GUI mode.")
                sys.exit(1)
            else:
                logger.debug("Using provided request token.")

        try:
            self.access_token = kite.generate_session(
                request_token, self.api_secret)['access_token']
            logger.debug("Access token generated successfully.")
            with open(token_file, "w") as file:
                file.write(str(self.access_token))
            logger.debug("Access token saved successfully.")
        except Exception as e:
            logger.error(f"Error generating access token: {e}")
            sys.exit(1)
        #return self.access_token

    def init_kite(self):
        self.kite = KiteConnect(api_key=self.api_key)
        self.kite.set_access_token(self.access_token)
        logger.debug("Kite Connect initialized successfully.")

    def get_api_key(self):
        return self.api_key
    
    def get_api_secret(self):
        return self.api_secret

    def get_access_token(self):
        return self.access_token