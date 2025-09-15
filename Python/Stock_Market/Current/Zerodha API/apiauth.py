import customtkinter as ctk
import webbrowser
from kitelogin import KiteLogin
from kiteconnect import KiteConnect
import tkinter.messagebox as tkmb
from baselogger import logger

class APIAuth():
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        #ctk.set_default_color_theme("blue")
        self.app = ctk.CTk()
        self.app.geometry("550x550")
        self.app.title("API Authentication")
        self.app.eval('tk::PlaceWindow . center')
        self.frame = ctk.CTkFrame(master=self.app)
        self.frame.pack(pady=20,padx=40, fill='both',expand=True)

        self.api_key = None
        self.api_secret = None
        self.access_token = None
        self.kite_login = KiteLogin()
        self.kite_login.load_credentials()

        self.api_key = self.kite_login.get_api_key()
        self.api_secret = self.kite_login.get_api_secret()

        if self.api_key and self.api_secret:
            # Already have credentials → Skip straight to URL/token step
            logger.debug("API Key and API Secret found.")
            self.request_token_screen()
        else:
            # First-time setup → Ask user for API Key and Secret
            logger.debug("No API Key and API Secret found, asking user to input them.")
            self.label = ctk.CTkLabel(master=self.frame, text="Enter API Key and API Secret")
            self.label.pack(pady=10)

            self.api_key_entry = ctk.CTkEntry(master=self.frame, placeholder_text="API Key", width=250)
            self.api_key_entry.pack(pady=10)

            self.api_secret_entry = ctk.CTkEntry(master=self.frame, placeholder_text="API Secret",width=250)
            self.api_secret_entry.pack(pady=10)

            self.submit_button2 = ctk.CTkButton(master=self.frame, text="Submit", command=self.save_key_and_secret)
            self.submit_button2.pack(pady=20)
        
        self.app.mainloop()

    def save_key_and_secret(self):
        key = self.api_key_entry.get()
        secret = self.api_secret_entry.get()
        if not key or not secret:
            tkmb.showerror(title="Error",message="Both API Key and API Secret are required")
            return
        self.kite_login.create_credentials(mode="gui", api_key=key, api_secret=secret)
        self.api_key = self.kite_login.get_api_key()
        self.api_secret = self.kite_login.get_api_secret()
        if not self.api_key or not self.api_secret:
            tkmb.showerror(title="Error",message="Failed to save API Key and API Secret. Please try again.")
            return
        tkmb.showinfo(title="Success",message="API Key and API Secret saved successfully. Proceeding to next step")
        # Clear initial widgets
        self.label.destroy()
        self.api_key_entry.destroy()
        self.api_secret_entry.destroy()
        self.submit_button2.destroy()
        
        # Show URL/token screen
        self.request_token_screen()

    def save_access_token(self):
        self.kite_login.load_access_token()
        self.access_token = self.kite_login.get_access_token()
        if not self.access_token:
            # No access token → Need to generate one
            logger.debug("No access token found, generating a new one.")
            token = self.token_entry.get()
            if not token:
                tkmb.showerror(title="Error",message="Kindly Input Request Token")
                return
            self.kite_login.create_access_token(mode="gui", request_token=token)
            self.access_token = self.kite_login.get_access_token()
            if not self.access_token:
                tkmb.showerror(title="Error",message="Failed to generate Access Token. Please try again.")
                return
            msg = "Access Token Generated and saved successfully"
        else:
            msg = "Access Token already exists, proceeding further"

        logger.debug(msg)
        tkmb.showinfo(title="Info",message=msg)
        #tkmb.showinfo(title="Success",message="All good, proceeding to Login")
        self.app.destroy()

    def request_token_screen(self):
        kite = KiteConnect(api_key=self.api_key)
        logger.debug(f"Login URL: {kite.login_url()}")
        login_url = kite.login_url()
        self.url_label1 = ctk.CTkLabel(master=self.frame, text="Open below URL in browser.Then copy and paste the generated Request Token from browser in Text box, and click Submit to generate Access Token", wraplength=450)
        self.url_label1.pack(pady=10)
        self.url_label2 = ctk.CTkEntry(master=self.frame, textvariable=ctk.StringVar(value=login_url), state="readonly", border_width=0, width=450)
        self.url_label2.pack(pady=10)

        #open_btn = ctk.CTkButton(master=self.frame, text="Open in Browser", command=lambda: webbrowser.open(kite.login_url()))
        #open_btn.pack()

        self.token_entry = ctk.CTkEntry(master=self.frame, placeholder_text="Paste Request Token", width=250)
        self.token_entry.pack(pady=10)

        self.submit_button1 = ctk.CTkButton(master=self.frame, text="Submit", command=self.save_access_token)
        self.submit_button1.pack(pady=20)