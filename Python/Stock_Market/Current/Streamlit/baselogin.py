# baselogin.py
import hashlib
import basedb
from baselogger import logger
from users import Users
import platform

def hash_password(password: str) -> str:
    salt = "UOHsUstBs+W+DleuSQ8vSw=="
    salted_pwd = password + salt
    """Hashes the plain password string using SHA-256 to match database standard."""
    return hashlib.sha256(salted_pwd.encode()).hexdigest()

def verify_user_login(username, password):
    """
    Validates user credentials.
    Returns True if valid, False otherwise.
    """
    login_flag = False
    logger.debug(f"Verifying login for user: {username}")
    data = Users.user_data
    for item in data["Users"]:
        if item["UserName"] == username and item["SystemInfo"]["NodeName"] == platform.uname().node:
            logger.debug("Node Name matches")
            login_flag = True
            break
    if not login_flag:
        logger.debug("Node Name does not match.")
        return False        
    stored_hash = basedb.get_stored_password_hash(username)
    if stored_hash is None:
        return False
    # Hash incoming text input and match with stored signature
    return stored_hash == hash_password(password)

def do_register(username, password):
    """
    Hashes the credentials and calls the database engine layer to create a profile.
    Returns True if registration succeeds, raises exception if it fails.
    """
    logger.debug(f"Attempting to register user: {username}")
    hashed_pwd = hash_password(password)
    success = basedb.insert_new_user(username, hashed_pwd)
    
    if not success:
        logger.error(f"Failed to register user: {username}. Username may already exist.")
        raise Exception("Username already exists or database execution error.")
    return True

def valid_user(user_name):
    logger.debug("Checking if valid user from the predefined list of users")
    valid = False
    data = Users.user_data
    for item in data["Users"]:
        if user_name == item["UserName"]:
            valid = True
            break
    return valid
