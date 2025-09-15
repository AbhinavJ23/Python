import hashlib
import platform
from baselogger import logger
from basedb import DBUtils
from users import Users
import customtkinter as ctk
import tkinter.messagebox as tkmb
from datetime import datetime
from nseindex import NseIndex

class Login:
    
    def __init__(self):
        super().__init__()
        self.is_logged_in = False
        self.index_symbol = None
        ctk.set_appearance_mode("dark")
        self.app = ctk.CTk()
        self.app.geometry("550x550")
        self.app.title("Application Login")
        self.app.eval('tk::PlaceWindow . center')

        self.frame = ctk.CTkFrame(master=self.app)
        self.frame.pack(pady=20,padx=40, fill='both',expand=True)

        self.user_id_label = ctk.CTkLabel(master=self.frame, text="User Id", justify="left")
        #self.user_id_label.pack(pady=12,padx=10)
        self.user_id_label.pack(pady=(10,0))
        self.user_id = ctk.CTkEntry(master=self.frame, placeholder_text="Enter User Id")
        #self.user_id.pack(pady=12,padx=10)
        self.user_id.pack(pady=0)

        self.user_pass_label = ctk.CTkLabel(master=self.frame, text="Password")
        #self.user_pass_label.pack(pady=12,padx=10)
        self.user_pass_label.pack(pady=(10,0))
        self.user_pass = ctk.CTkEntry(master=self.frame, placeholder_text="Enter Password", show="*")
        #self.user_pass.pack(pady=12,padx=10)
        self.user_pass.pack(pady=0)

        self.index = NseIndex()
        self.options = self.index.index_symbols
        self.index_label = ctk.CTkLabel(master=self.frame, text="Index")
        #self.index_label.pack(pady=12,padx=10)
        self.index_label.pack(pady=(10,0))
        self.index_combo = ctk.CTkComboBox(master=self.frame, values=self.options)
        #self.index_combo.pack(pady=12,padx=10)
        self.index_combo.pack(pady=0)
        self.index_combo.set("NIFTY 50")

        #self.index_combo.pack(pady=12)
        self.button2 = ctk.CTkButton(master=self.frame, text="Register", command=self.do_register)
        #self.button2.pack(pady=(10,0))
        self.button2.pack(padx=12, pady=(25,0))

        self.button1 = ctk.CTkButton(master=self.frame, text="Login", command=self.do_login)
        self.button1.pack(padx=12, pady=10)

        self.button3 = ctk.CTkButton(master=self.frame, text="Clear", command=self.do_clear)
        self.button3.pack(padx=12, pady=0)

        self.system_info = platform.uname()
        self.system = self.system_info.system
        self.node = self.system_info.node

        self.db = DBUtils()
        try:
            self.connection = self.db.create_connection('login.db')
            self.cursor = self.connection.cursor()
        except Exception as e:
            logger.error(f'Error creating database - {e}')
        
        self.app.mainloop()

    def do_clear(self):
        logger.debug("do_clear")
        self.user_id.delete(0, ctk.END)
        self.user_pass.delete(0, ctk.END)        

    def hash_password(self, pwd):
        salt = "UOHsUstBs+W+DleuSQ8vSw=="
        salted_pwd = pwd + salt
        hashed_pwd = hashlib.sha256(salted_pwd.encode()).hexdigest()
        return hashed_pwd

    def valid_user(self, user_id):
        logger.debug("Checking valid user")
        valid = False
        data = Users.user_data
        for item in data["Users"]:
            if user_id == item["UserId"]:
                valid = True
                break
        return valid
    
    def empty_user_pass(self, user_str, password_str):
        flag = False
        if not user_str or not password_str:
            logger.debug("User Id/Password is Null")
            flag = True
        return flag
    
    def log_user_login(self, user_str):
        date_str = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        user_ins_query = """INSERT INTO USER_LOGON_DATA (USER_DATA_ID,DATE) VALUES (?,?);"""
        try:
            self.cursor.execute(user_ins_query, (user_str,date_str))
            self.connection.commit()
            logger.debug("User details logged successfully")
            return True
        except Exception as e:
            logger.error(f'Error logging user details - {e}')
            self.do_clear()
            return False
        
    def do_cleanup(self):
        self.cursor.close()
        self.connection.close()
        self.app.destroy()
    
    def do_register(self):
        register_flag = True
        user_str = self.user_id.get()
        password_str = self.user_pass.get()
        logger.debug(f'Registering - {user_str}')
        
        if self.empty_user_pass(user_str, password_str):
            logger.debug("do_register - User Id/Password is Null")
            tkmb.showerror(title="Error",message="Please enter User Id/Password")
            self.do_clear()
            register_flag = False
        else:            
            valid = self.valid_user(user_str)
            if valid:
                logger.debug("do_register - User is Valid")
                user_query = """
                    CREATE TABLE IF NOT EXISTS USER_DATA (
                    USER_ID TEXT PRIMARY KEY,
                    DATE TEXT NOT NULL
                    );
                    """
                pwd_query = """
                    CREATE TABLE IF NOT EXISTS USER_PASSWORD (
                    USER_DATA_ID TEXT PRIMARY KEY,
                    USER_PWD TEXT NOT NULL,
                    DATE TEXT NOT NULL,
                    FOREIGN KEY (USER_DATA_ID) REFERENCES USER_DATA (USER_ID)
                    );
                    """                
                user_logon_query = """
                    CREATE TABLE IF NOT EXISTS USER_LOGON_DATA (
                    USER_DATA_ID TEXT NOT NULL,
                    DATE TEXT NOT NULL,
                    FOREIGN KEY (USER_DATA_ID) REFERENCES USER_DATA (USER_ID)
                    );
                    """
                query = user_query
                status = self.db.execute_query(self.connection, user_query)
                if status:
                    query = pwd_query
                    status = self.db.execute_query(self.connection, pwd_query)                    
                if status:
                    query = user_logon_query
                    status = self.db.execute_query(self.connection, user_logon_query)                    

                if not status:
                    logger.error(f'Error creating table for query - {query}')
                    register_flag = False
                else:
                    self.connection.commit()
                    hashed_pwd = self.hash_password(password_str)
                    date_str = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
                    user_ins_query = """INSERT INTO USER_DATA (USER_ID,DATE) VALUES (?,?);"""
                    pass_ins_query = """INSERT INTO USER_PASSWORD (USER_DATA_ID,USER_PWD,DATE) VALUES (?,?,?);"""
                    try:
                        self.cursor.execute(user_ins_query, (user_str,date_str))
                        self.cursor.execute(pass_ins_query, (user_str,hashed_pwd,date_str))
                        self.connection.commit()
                        logger.debug("User Id and Password inserted successfully")
                        tkmb.showinfo(title="Success",message="User registered Successfully")                    
                        self.do_clear()
                    except Exception as e:
                        logger.error(f'Error inserting User Id/Password - {e}')
                        tkmb.showerror(title="Error",message="Error registering User")
                        self.do_clear()
                        register_flag = False
            else:
                logger.debug("do_register - User Id is Invalid")
                tkmb.showerror(title="Error",message="Please enter valid User Id")
                self.do_clear()
                register_flag = False
        return register_flag

    def do_login(self):
        logger.debug("do_login")
        data = Users.user_data
        logger.debug(f'System - {self.system}')        
        input_user_id = self.user_id.get()
        input_password = self.user_pass.get()
        if self.empty_user_pass(input_user_id, input_password):
            logger.debug("do_login - User Id/Password is Null")
            tkmb.showerror(title="Error",message="Please enter User Id/Password")
            self.do_clear()
            return False
        else:
            if not self.valid_user(input_user_id):
                logger.debug("do_login - User is Invalid")
                tkmb.showerror(title="Error",message="Please enter valid User Id")
                self.do_clear()
                return False
            elif self.index_combo.get() not in self.index.index_symbols:
                logger.debug("do_login - Invalid Index selected")
                tkmb.showerror(title="Error",message="Please select valid Index")
                self.do_clear()
                return False
            else:
                logger.debug("do_login - User is Valid")

        table_name = 'USER_DATA'
        table_flag = self.db.table_exists(self.connection, table_name)
        if not table_flag:            
            tkmb.showerror(title="Error",message="Please check if you have registerd")
            self.do_clear()
            return False
        
        db_password = ""
        select_pwd_query = """SELECT USER_PWD FROM USER_PASSWORD WHERE USER_DATA_ID = ?;"""
        try:
            self.cursor.execute(select_pwd_query, (input_user_id,))
            record = self.cursor.fetchone()
            if record:
                db_password = record[0]
            else:
                logger.debug("User Id does not exist, register first")
                tkmb.showerror(title="Error",message="Please check if you have registerd")
                self.do_clear()
                return False
        except Exception as e:
            logger.error(f'Error selecting User Id - {e}')
            tkmb.showerror(title="Error",message="Exception Occurred!")
            self.do_clear()
            return False

        login_flag = False
        if self.hash_password(input_password) == db_password:
            logger.debug("Password matches")
            for item in data["Users"]:
                if item["UserId"] == input_user_id and item["SystemInfo"]["NodeName"] == self.node:
                    logger.debug("Node Name matches")
                    tkmb.showinfo(title="Login Successful",message="You have logged in Successfully")
                    status = self.log_user_login(input_user_id)
                    if not status:
                        tkmb.showerror(title="Error",message="Failed to insert user logon details")
                    self.do_clear()
                    login_flag = True
                    self.is_logged_in = True
                    self.index_symbol = self.index_combo.get()
                    logger.debug(f"Selected Index - {self.index_symbol}")
                    self.do_cleanup()
                    break
            if not login_flag:
                logger.debug("Node Name does not match")
                tkmb.showerror(title="Login Failed",message="System Info does not match")
                self.do_clear()
                return False
        else:
            logger.debug("Password does not match")
            tkmb.showerror(title="Login Failed",message="Invalid User Id/password")
            self.do_clear()

        return login_flag