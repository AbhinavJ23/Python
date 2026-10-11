import json
import os
from baselogger import logger
from kiteconnect import KiteConnect

credentials_file="credentials.json"

def save_api_credentials(api_key, api_secret):
    """Saves API Key and Secret to a local JSON file."""
    #os.makedirs(os.path.dirname(credentials_file), exist_ok=True)
    with open(credentials_file, "w") as f:
        json.dump({"api_key": api_key, "api_secret": api_secret}, f)
    logger.debug("API Credentials saved successfully.")

def load_api_credentials():
    """Loads API Key and Secret if they exist."""
    if os.path.exists(credentials_file):
        with open(credentials_file, "r") as f:
            return json.load(f)
        logger.debug("API Credentials loaded successfully.")
    return None

def get_zerodha_login_url(api_key):
    """Generates the daily request token login URL."""
    kite = KiteConnect(api_key=api_key)
    return kite.login_url()

def generate_session_token(request_token, api_key, api_secret):
    """Exchanges request token for access token and returns the kite object."""
    kite = KiteConnect(api_key=api_key)
    data = kite.generate_session(request_token, api_secret=api_secret)
    # Return both the access token and the initialized kite instance
    return data["access_token"], kite