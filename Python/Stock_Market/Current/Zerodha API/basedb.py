import sqlite3
from baselogger import logger

class DBUtils:

    def __init__(self):
        super().__init__()

    def create_connection(self, path):
        connection = None
        try:
            connection = sqlite3.connect(path)
            logger.debug("Connection to SQLite DB successful")
            return connection
        except Exception as e:
            logger.error(f'Error creating connection - {e}')
            return None
    
    def execute_query(self, connection, query):
        cursor = connection.cursor()
        try:
            cursor.execute(query)
            logger.debug("Query executed successfully")
            return True
        except Exception as e:
            logger.error(f'Error executing query - {e}')
            return False
        
    def table_exists(self, connection, table_name):
        cursor = connection.cursor()
        select_table_query = """SELECT name FROM sqlite_master WHERE type='table' AND name = ?;"""
        try:
            cursor.execute(select_table_query, (table_name,))
            record = cursor.fetchall()
            if record == []:
                logger.debug(f'Table - {table_name} - does not exist')
                return False
            else:
                logger.debug(f'Table - {table_name} - exists')
                return True
        except Exception as e:
            logger.error(f'Error executing query - {e}')
            return False


