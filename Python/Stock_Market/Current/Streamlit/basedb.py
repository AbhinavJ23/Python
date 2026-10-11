import sqlite3
import os
from baselogger import logger
from datetime import datetime

DB_NAME = "app_data.db"  # Stores database in your execution folder

def create_connection():
    """Establishes and returns an active SQLite connection handle."""
    connection = None
    try:
        connection = sqlite3.connect(DB_NAME)
        logger.debug("Connection to SQLite DB successful")
        return connection
    except Exception as e:
        logger.error(f'Error creating connection - {e}')
        return None
    
def execute_query(connection, query):
    cursor = connection.cursor()
    try:
        cursor.execute(query)
        logger.debug("Query executed successfully")
        return True
    except Exception as e:
        logger.error(f'Error executing query - {e}')
        return False

def setup_database():
    """
    Creates necessary database tables if they do not exist.
    Runs once when the application boots up.
    """
    conn = create_connection()
    #cursor = conn.cursor()
    
    # Create the user authentication profiles matrix table
    user_query = """
        CREATE TABLE IF NOT EXISTS USER_DATA (
            USER_ID INTEGER PRIMARY KEY AUTOINCREMENT,
            USER_NAME TEXT UNIQUE NOT NULL,
            PASSWORD TEXT NOT NULL,
            CREATED_AT TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        """
    user_logon_query = """
        CREATE TABLE IF NOT EXISTS USER_LOGON_DATA (
            USER_NAME TEXT NOT NULL,
            LOGGED_AT TEXT NOT NULL
            );
        """
    #FOREIGN KEY (USER_DATA_ID) REFERENCES USER_DATA (USER_ID)
    query = user_query
    status = execute_query(conn, query)
    if status:
        query = user_logon_query
        status = execute_query(conn, query)
    if not status:
        logger.error(f'Error creating table for query - {query}')
        return False
    
    conn.commit()
    conn.close()

def get_stored_password_hash(username):
    """
    Queries database for a given username.
    Returns the string hash value if found, otherwise returns None.
    """
    conn = create_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT PASSWORD FROM USER_DATA WHERE USER_NAME = ?", (username,))
        row = cursor.fetchone()
        # row is a tuple, e.g., ("hashed_string_here",)
        return row[0] if row else None
    except sqlite3.OperationalError:
        return None
    finally:
        conn.close()

def insert_new_user(username, password):
    """Inserts a new user profile record into the database."""
    logger.debug(f'Registering - {username}')
    conn = create_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO USER_DATA (USER_NAME, PASSWORD) VALUES (?, ?)", (username, password))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        # Username already exists
        return False
    finally:
        conn.close()

def log_user_login(user_name):
    logger.debug(f'Logging in - {user_name}')
    conn = create_connection()
    cursor = conn.cursor()
    date_str = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    user_ins_query = """INSERT INTO USER_LOGON_DATA (USER_NAME,LOGGED_AT) VALUES (?,?);"""
    try:
        cursor.execute(user_ins_query, (user_name, date_str))
        conn.commit()
        logger.debug("User details logged successfully")
        return True
    except Exception as e:
        logger.error(f'Error logging user details - {e}')
        return False