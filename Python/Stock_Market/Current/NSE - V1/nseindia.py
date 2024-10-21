import requests
import pandas as pd
from baselogger import logger
from http import HTTPStatus

pd.set_option('display.max_rows', 1000)
pd.set_option('display.max_columns', 1000)
pd.set_option('display.width', 5000)

class NSE:
    pre_market_categories = ['NIFTY 50', 'Nifty Bank', 'Emerge', 'Securities in F&O', 'Others', 'All' ]
    equity_market_categories = ['NIFTY 50', 'NIFTY NEXT 50', 'NIFTY 100', 'NIFTY 200', 'NIFTY 500', 'NIFTY MIDCAP 50', 'NIFTY MIDCAP 100', 'NIFTY SMALLCAP 100', 'INDIA VIX', 'NIFTY MIDCAP 150', 'NIFTY SMALLCAP 50', 'NIFTY SMALLCAP 250', 'NIFTY MIDSMALLCAP 400', 'NIFTY500 MULTICAP 50:25:25', 'NIFTY LARGEMIDCAP 250', 'NIFTY MIDCAP SELECT', 'NIFTY TOTAL MARKET', 'NIFTY MICROCAP 250', 'NIFTY BANK', 'NIFTY AUTO', 'NIFTY FINANCIAL SERVICES', 'NIFTY FINANCIAL SERVICES 25/50', 'NIFTY FMCG', 'NIFTY IT', 'NIFTY MEDIA', 'NIFTY METAL', 'NIFTY PHARMA', 'NIFTY PSU BANK', 'NIFTY PRIVATE BANK', 'NIFTY REALTY', 'NIFTY HEALTHCARE INDEX', 'NIFTY CONSUMER DURABLES', 'NIFTY OIL & GAS']
    holiday_categories = ['Clearing', 'Trading']

    def __init__(self):
        self.headers = {#"accept-encoding":"gzip, deflate, br, zstd",
                        "accept-language":"en-US,en;q=0.9",
                        "user-agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                                    "Chrome/129.0.0.0 Safari/537.36"}
        
        ##### FOR FUTURE USE IF REQUIRED #####
        #self.headers = {"accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,/;q=0.8,application/signed-exchange;v=b3;q=0.7",
        #                "accept-language": "en-US,en;q=0.9,en-IN;q=0.8,en-GB;q=0.7",
        #                "cache-control": "max-age=0",
        #                "priority": "u=0, i",
        #                "sec-ch-ua": """Google Chrome";v="129", "Not=A?Brand";v="8", "Chromium";v="129""",
        #                "sec-ch-ua-mobile": "?0",
        #                "sec-ch-ua-platform": "Windows",
        #                "sec-fetch-dest": "document",
        #                "sec-fetch-mode": "navigate",
        #                "sec-fetch-site": "none",
        #                "sec-fetch-user": "?1",
        #                "upgrade-insecure-requests": "1",
        #                "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36 Edg/129.0.0.0"}
        
        try:
            self.session = requests.Session()
            self.session.get('https://www.nseindia.com', headers= self.headers)
        except (requests.exceptions.ConnectionError) as e:
            logger.error(f'Function __init__ - Error - {e}')

    def pre_market_data(self, category):
        pre_market_category = {"NIFTY 50": "NIFTY", "Nifty Bank": "BANKNIFTY", "Emerge": "SME", "Securities in F&O":"FO", 
                            "Others": "OTHERS", "All": "ALL"}
        try:
            response = self.session.get(f'https://www.nseindia.com/api/market-data-pre-open?key={pre_market_category[category]}', headers=self.headers)
            response.raise_for_status()
        except (requests.HTTPError, requests.exceptions.ConnectionError, requests.exceptions.ChunkedEncodingError, 
                requests.ConnectionError, requests.exceptions.HTTPError) as e:
            logger.error(f'Function pre_market_data - Error - {e}')
            return None
        if response.status_code != 401:
            try:
                data = response.json()["data"]
                temp_data = []
                for i in data:
                    temp_data.append(i["metadata"])
                df = pd.DataFrame(temp_data)
                df = df.set_index('symbol', drop=True)
                return df
            except (requests.JSONDecodeError,requests.exceptions.JSONDecodeError,ValueError,KeyError) as e:
                logger.error(f'Function pre_market_data - Decoding JSON has failed - {e}')
                return None
    
    def equity_market_data(self, category, symbol_list=False):
        category = category.upper().replace(' ', '%20').replace('&', '%26')
        try:
            response = self.session.get(f'https://www.nseindia.com/api/equity-stockindices?index={category}', headers=self.headers)
            response.raise_for_status()
        except (requests.HTTPError, requests.exceptions.ConnectionError, requests.exceptions.ChunkedEncodingError, 
                requests.ConnectionError, requests.exceptions.HTTPError) as e:
            logger.error(f'Function equity_market_data - Error - {e}')
            return None
        if response.status_code != 401:
            try:
                data = response.json()["data"]
                df = pd.DataFrame(data)
                df = df.drop("meta", axis=1)
                df = df.set_index('symbol', drop=True)
                if symbol_list:
                    return list(df.index)
                else:
                    return df
            except (requests.JSONDecodeError,requests.exceptions.JSONDecodeError,ValueError,KeyError) as e:
                logger.error(f'Function equity_market_data - Decoding JSON has failed - {e}')
                return None
        else:
            return None
        
    def top_gainers_loosers(self, gainers = True):
        category = "Securities in F&O"
        category = category.upper().replace(' ', '%20').replace('&', '%26')
        try:
            response = self.session.get(f'https://www.nseindia.com/api/equity-stockindices?index={category}', headers=self.headers)
            response.raise_for_status()
        except (requests.HTTPError, requests.exceptions.ConnectionError, requests.exceptions.ChunkedEncodingError, 
                requests.ConnectionError, requests.exceptions.HTTPError) as e:
            logger.error(f'Function top_gainers_loosers - Error - {e}')
            return None
        if response.status_code != 401:
            try:
                data = response.json()["data"]
                df = pd.DataFrame(data)
                if gainers:
                    df.sort_values(by="pChange", ascending = False)
                else:
                    df = df.sort_values(by="pChange")
                return df.head(5)
            except (requests.JSONDecodeError,requests.exceptions.JSONDecodeError,ValueError,KeyError) as e:
                logger.error(f'Function top_gainers_loosers - Decoding JSON has failed - {e}')
                return None
        
    def nse_holidays(self, type):
        try:
            response = self.session.get(f'https://www.nseindia.com/api/holiday-master?type={type.lower()}', headers=self.headers)
            response.raise_for_status()
        except (requests.HTTPError, requests.exceptions.ConnectionError, requests.exceptions.ChunkedEncodingError, 
                requests.ConnectionError, requests.exceptions.HTTPError) as e:
            logger.error(f'Function nse_holidays - Error - {e}')
            return None
        if response.status_code != 401:
            try:
                data = response.json()
                df = pd.DataFrame(list(data.values())[0])
                return df
            except (requests.JSONDecodeError,requests.exceptions.JSONDecodeError,ValueError,KeyError) as e:
                logger.error(f'Function nse_holidays - Decoding JSON has failed - {e}')
                return None
    
    def equity_info(self, symbol, trade_info=False):        
        symbol = symbol.replace(' ', '%20').replace('&', '%26')
        try:
            response = self.session.get("https://www.nseindia.com/api/quote-equity?symbol=" + symbol + ("&section=trade_info" if trade_info else ""), headers=self.headers)
            response.raise_for_status()
        except (requests.HTTPError, requests.exceptions.ConnectionError, requests.exceptions.ChunkedEncodingError, 
                requests.ConnectionError, requests.exceptions.HTTPError) as e:
            logger.error(f'Function equity_info - Error - {e}')
            return None
        if response.status_code != 401:
            try:
                data = response.json()
                return data
            except (requests.JSONDecodeError,requests.exceptions.JSONDecodeError,ValueError,KeyError) as e:
                logger.error(f'Function equity_info - Decoding JSON has failed - {e}')
                return None
        else:
            return None            
    
    def futures_data(self, symbol, indices=False):
        symbol = symbol.replace(' ', '%20').replace('&', '%26')
        try:
            response = self.session.get("https://www.nseindia.com/api/quote-derivative?symbol=" + symbol, headers=self.headers)
            response.raise_for_status()
        except (requests.HTTPError, requests.exceptions.ConnectionError, requests.exceptions.ChunkedEncodingError, 
                requests.ConnectionError, requests.exceptions.HTTPError) as e:
            logger.error(f'Function futures_data - Error - {e}')
            return None
        if response.status_code != 401:
            try:
                data = response.json()
                temp_data = []
                for i in data["stocks"]:
                    if i["metadata"]["instrumentType"] == ("Index Futures" if indices else "Stock Futures"):
                        temp_data.append(i["metadata"])

                df = pd.DataFrame(temp_data)
                df = df.set_index("identifier", drop=True)
                return df
            except (requests.JSONDecodeError,requests.exceptions.JSONDecodeError,ValueError,KeyError) as e:
                logger.error(f'Function futures_data - Decoding JSON has failed - {e}')
                return None
        else:
            return None

    def derivatives_data(self, symbol):
        symbol = symbol.replace(' ', '%20').replace('&', '%26')
        try:
            response = self.session.get("https://www.nseindia.com/api/quote-derivative?symbol=" + symbol, headers=self.headers)
            response.raise_for_status()
        except (requests.HTTPError, requests.exceptions.ConnectionError, requests.exceptions.ChunkedEncodingError, 
                requests.ConnectionError, requests.exceptions.HTTPError) as e:
            logger.error(f'Function derivatives_data - Error - {e}')
            return None
        if response.status_code != 401:
            try:
                data = response.json()
                return data
            except (requests.JSONDecodeError,requests.exceptions.JSONDecodeError,ValueError,KeyError) as e:
                logger.error(f'Function  derivatives_data - Decoding JSON has failed - {e}')
                return None
        else:
            return None
            
    def options_data(self, symbol, indices=False):
        symbol = symbol.replace(' ', '%20').replace('&', '%26')
        if not indices:
            url = "https://www.nseindia.com/api/option-chain-equities?symbol="+ symbol
        else:
            url = "https://www.nseindia.com/api/option-chain-indices?symbol="+ symbol
        try:
            response = self.session.get(url, headers=self.headers)
            response.raise_for_status()
        except (requests.HTTPError, requests.exceptions.ConnectionError, requests.exceptions.ChunkedEncodingError, 
                requests.ConnectionError, requests.exceptions.HTTPError) as e:
            logger.error(f'Function options_data - Error - {e}')
            return None        
        if response.status_code != 401:
            try:
                data = response.json()["records"]
                op_data = []
                for i in data["data"]:
                    for key,value in i.items():
                        if key == "CE" or key == "PE":
                            info = value
                            info["instrumentType"] = key
                            info["timestamp"] = data["timestamp"]
                            op_data.append(info)
                df = pd.DataFrame(op_data)
                df = df.set_index("identifier", drop=True)
                return df
            except (requests.JSONDecodeError,requests.exceptions.JSONDecodeError,ValueError,KeyError) as e:
                logger.error(f'Function options_data - Decoding JSON has failed - {e}')
                return None
        else:
            return None   