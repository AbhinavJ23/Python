from kitelogin import KiteLogin
from kiteconnect import KiteConnect
import pandas as pd
import datetime
import xlwings as xw
from baselogger import logger
import os,sys,time
from nseindex import NseIndex

class SupportResistance:
    def __init__(self):
        self.kite_login = KiteLogin()
        self.api_key = self.kite_login.get_api_key()
        self.access_token = self.kite_login.get_access_token()
        self.kite = KiteConnect(api_key=self.api_key)
        self.kite.set_access_token(self.access_token)
        
        self.index = NseIndex()
        self.equity_market_categories = self.index.index_symbols
        self.nifty50_symbols = self.index.nifty_50_symbols
        self.niftynext_50_symbols = self.index.nifty_next_50_symbols
        self.niftymidcap_50_symbols = self.index.nifty_midcap_50_symbols
        self.niftysmallcap_50_symbols = self.index.nifty_smallcap_50_symbols

    def calculate_support_resistance(self, index_symbol):

        symbols = []
        if index_symbol == 'NIFTY 50':
            symbols = self.nifty50_symbols
        elif index_symbol == 'NIFTY NEXT 50':
            symbols = self.niftynext_50_symbols
        elif index_symbol == 'NIFTY MIDCAP 50':
            symbols = self.niftymidcap_50_symbols
        elif index_symbol == 'NIFTY SMALLCAP 50':
            symbols = self.niftysmallcap_50_symbols
        else:
            logger.error(f"Support Resistance for {index_symbol} is not implemented.")
            return None

        """ NIFTY 50 stock symbols
        nifty_50_symols = ["ADANIENT", "ADANIPORTS", "APOLLOHOSP", "ASIANPAINT", "AXISBANK", "BAJAJ-AUTO", 
                    "BAJFINANCE", "BAJAJFINSV", "BEL", "BHARTIARTL", 
                    "CIPLA", "COALINDIA", "DRREDDY", "EICHERMOT", "ETERNAL", 
                    "GRASIM", "HCLTECH", "HDFCBANK", "HDFCLIFE", "HEROMOTOCO", 
                    "HINDALCO", "HINDUNILVR", "ICICIBANK", "INDUSINDBK", 
                    "INFY", "ITC", "JIOFIN", "JSWSTEEL", "KOTAKBANK", "LT", 
                    "M&M", "MARUTI", "NESTLEIND", "NTPC", "ONGC", 
                    "POWERGRID", "RELIANCE", "SBILIFE", "SBIN", "SHRIRAMFIN", "SUNPHARMA", 
                    "TATACONSUM", "TATAMOTORS", "TATASTEEL", "TCS", "TECHM", 
                    "TITAN", "TRENT", "ULTRACEMCO", "WIPRO"
        ]
        """

        # Get instrument tokens
        instruments = self.kite.instruments("NSE")
        nse_map = {i['tradingsymbol']: i['instrument_token'] for i in instruments if i['tradingsymbol'] in symbols}

        # Fetch OHLC and calculate Support and Resistance
        data = []
        today = datetime.date.today()
        from_date = today - datetime.timedelta(days=2)
        to_date = today - datetime.timedelta(days=1)

        # If to_date is Saturday or Sunday, adjust to_date to last Friday and from_date to last Thursday
        if to_date.weekday() >= 5:
            to_date = to_date - datetime.timedelta(days=to_date.weekday() - 4)
            from_date = from_date - datetime.timedelta(days=from_date.weekday() - 3)

        for symbol in symbols:
            try:
                token = nse_map[symbol]
                ohlc = self.kite.historical_data(token, from_date, to_date, "day")
                if ohlc:
                    latest = ohlc[-1]
                    high = latest['high']
                    low = latest['low']
                    close = latest['close']
                    pivot = (high + low + close) / 3
                    s1 = 2 * pivot - high
                    r1 = 2 * pivot - low
                    data.append([symbol, round(s1, 2), round(r1, 2)])
            except Exception as e:
                logger.error(f"Error occured for {symbol}: {e}")

        # Lets Format output
        df = pd.DataFrame(data, columns=["Stock", "Support", "Resistance"])
        df.set_index("Stock", inplace=True)
        df = df.transpose()  # Support is first row, Resistance second
        return df

    def save_to_excel(self, df, index_symbol):
        ## Creating new excel and adding sheets
        file_name = 'SupportResistance_'+time.strftime('%Y%m%d%H%M%S')+'.xlsx'
        if not os.path.exists(file_name):
            try:
                wb = xw.Book()
                wb.sheets.add(index_symbol)
                wb.save(file_name)
                logger.debug("Created Excel - " + file_name)
            except Exception as e:
                logger.critical(f'Error Creating Excel - {e}')
                sys.exit()
        wb = xw.Book(file_name)
        sr = wb.sheets(index_symbol)
        sr.range("A1").value = df
        logger.debug(f"Support and Resistance data for {index_symbol} saved to Excel file.")