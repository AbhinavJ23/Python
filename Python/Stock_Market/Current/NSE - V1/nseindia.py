import requests
import pandas as pd
import requests.cookies
from baselogger import logger
from http import HTTPStatus
import json

pd.set_option('display.max_rows', 1000)
pd.set_option('display.max_columns', 1000)
pd.set_option('display.width', 5000)

class NSE:
    pre_market_categories = ['NIFTY 50', 'Nifty Bank', 'Emerge', 'Securities in F&O', 'Others', 'All' ]
    equity_market_categories = ['NIFTY 50', 'NIFTY NEXT 50', 'NIFTY 100', 'NIFTY 200', 'NIFTY 500', 'NIFTY MIDCAP 50', 'NIFTY MIDCAP 100', 'NIFTY SMALLCAP 100', 'INDIA VIX', 'NIFTY MIDCAP 150', 'NIFTY SMALLCAP 50', 'NIFTY SMALLCAP 250', 'NIFTY MIDSMALLCAP 400', 'NIFTY500 MULTICAP 50:25:25', 'NIFTY LARGEMIDCAP 250', 'NIFTY MIDCAP SELECT', 'NIFTY TOTAL MARKET', 'NIFTY MICROCAP 250', 'NIFTY BANK', 'NIFTY AUTO', 'NIFTY FINANCIAL SERVICES', 'NIFTY FINANCIAL SERVICES 25/50', 'NIFTY FMCG', 'NIFTY IT', 'NIFTY MEDIA', 'NIFTY METAL', 'NIFTY PHARMA', 'NIFTY PSU BANK', 'NIFTY PRIVATE BANK', 'NIFTY REALTY', 'NIFTY HEALTHCARE INDEX', 'NIFTY CONSUMER DURABLES', 'NIFTY OIL & GAS']
    holiday_categories = ['Clearing', 'Trading']

    def __init__(self):
        #self.headers = {"accept-encoding":"gzip, deflate, br, zstd",
        #                "accept-language":"en-US,en;q=0.9",
        #                "content-type":"application/json; charset=utf-8",
        #                "accept":"application/json,text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        #                "user-agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
        #                            "Chrome/129.0.0.0 Safari/537.36"}
        #self.cookies = '_ga=GA1.1.1009728837.1713533533; _ga_QJZ4447QD3=GS1.1.1716785431.25.1.1716786267.0.0.0; defaultLang=en; nseQuoteSymbols=[{"symbol":"ICICIBANK","identifier":null,"type":"equity"}]; _abck=D42845E7E775ABF23A53EC37A5250324~0~YAAQBdjIF7lT2miSAQAAgmFlkwwZ5daWA1PxfmLhfstN/lcURY02cTOLFOxzRQkg2JKUHLS0LO8YLUZb7zQsAEhCtFUuMiHZRoDO8EUoWfkcT0sk6nJjri13GegSUG91G513AxasA1v9Cos0LaXMljNYvIlebQxfvjhTyzZKE9EEh4CN7iLxSg6vWmgCwaHHhT8w0GtyDaK39WOIV0Fq0V2OPCndDhsCrh678Sg3vt15GnSOXp/YBIhhstnb6pUxL1Nh3/NlmB8p+Jw8CjFEM30HXJ28w0hJnmaR91SUovn28j7tiiYuCFlzo4aNjtx+6VTLo2of1OaHfsfnN508Q75t6UglpaCb+ShgKolvCfA2mVibSsn7mLj6jKOvEY3lgXptoJgGfjVeO2noLGuqc6+z7SAN182lxxo=~-1~-1~-1; nseappid=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJhcGkubnNlIiwiYXVkIjoiYXBpLm5zZSIsImlhdCI6MTcyOTA1NTY2MSwiZXhwIjoxNzI5MDYyODYxfQ.stFPY8naH6dxtKkhem_Q-4dRjtPhwfFDFpO3u2V3eh4; RT="z=1&dm=nseindia.com&si=7e680d20-fdd7-43e5-9dfd-d62d3cfdb1ae&ss=m2bcluo7&sl=1&se=8c&tt=0&bcn=%2F%2F684d0d43.akstat.io%2F&obo=1&ld=2hq9u&r=9w68x89&ul=2hqa1&hd=2hqa8"; _ga_87M7PJ3R97=GS1.1.1729055664.74.1.1729055664.60.0.0; AKA_A2=A; ak_bmsc=9D13F81A259BE7221A0FE345AC890717~000000000000000000000000000000~YAAQDdjIF4JKJ3KSAQAAdrTtkxnnT179m4mwA3ceiKy1gJBt+v+a1qSaMUzcsHkXWNgr2kdKNCMZZE1bdc77e6ELZmWpDXktYfVvmFhYqZUfJHrkJtJVExUI6Im7VyJL+nuDgTIEl87P+ACIAYHyG+OhVnZ8El2F5q28Agn8AZm9tf9dEi9OojSXqcNrC3LEeoJ7KVRQ3JAa5RUWKoH3/io6x0taJW4cRIwzfnMZSL3VXNysvsFcml9XHSI+/qn7RQFQl+V07+XiF1MvqbUcOC+EgoD6aT+DPkl6ZDhIx/CeEz/gbixw9bK+IkZ1pt++Z+U3ImQ0Td3KKnsuoJ9rS0ElgVtepuLVjsvGxALZPK0SudiN5Gh5okQyaUDkC+vA7MIVCgxvIOZ4VvU=; bm_sv=29E1D41E6030DFB4168D6FF77FD11FD1~YAAQ93UsMdpfnXGSAQAAvcf6kxkTqDMsEm31lWg/R9Gc5rGqYudMl4ecEt8PtJopYOSjaojpYaSDGo8DVcQA/7XPStAIhhSlLV9tX9WUTtdfl3o8LyGPtsjwot967pdtxhLGYyI4fZBt1p7/IEFGOpKXcyDLgvcQqDb83ki4/5P9tyaCTTtU6VtRu/VnqrWZAwaJkfubgAL9TIBM8tlaTgYOw0ALLQH5zRAxbWzmKtlVNZh6GbspAP/hN6mLooSMsRg=~1; bm_sz=32A01CC508FC106563FBFDB57AB2A185~YAAQ93UsMdtfnXGSAQAAvcf6kxlMd0HmeZk5NXlqxCSJIWn79TtwynkJ3mtyg6UxEMEG7GEAPl54P0Z9HcSUUtee7zVMtZ2y56APKA6lYsGwVjvcfTny41z4/+YdUMbvI4bynhPqFiaHVoerqwnafNQfcAEW/lx/4LAaUnfNV6/wAnyEr1doM4evo5FQaWm+OLnd8c1z4jmpVa/z3AXJLaaa1tpSdthRaOtJKgyfWAYP4GQUQ+pLZ3msQWmIjPfewnkuu8p1q3/2UFWF+G7hX7u0FgP3tQAaz/J8E/GFKuqiSUEk4+28lkG4zvnSTYu0TMeWBc5I2qDTY3RWAPNokPQ5lnQJD53NHql9br95AVorBG0ZD+9Ol5WFVfuFaSHXQh7XM2XO/mw9V2Tc4NU/qENlKspRbUZsCnkLKtsbBC0GMcNgdyr9oEIRuJGkIlHglJrUofuNnf3HwTbJ4CmgCgz7kjuKtwjlxeL8JspZ3redLD4XC7nw8CaZ/33sALpVHnd8Y+Df3XcEBu2hReCNeFRammi2zl5x4zkbDIeS7pc=~3486532~3687235'
        self.cookies = requests.cookies
        #self.cookies: Dict[str, str] = {}
        self.headers = {
            'Accept-Encoding':'gzip, deflate, br, zstd',
            'Accept':'application/json,text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
            'Accept-Language': 'en-US,en;q=0.9',
            'Connection': 'keep-alive',
            'Content-Type': 'application/json; charset=UTF-8',
            'Sec-Fetch-Dest': 'document',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-Site': 'none',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36',
            'X-Requested-With': 'XMLHttpRequest',
            'sec-ch-ua': '"Google Chrome";v="123", "Not:A-Brand";v="8", "Chromium";v="123"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
        }
        try:
            self.session = requests.Session()
            self.session.get('https://www.nseindia.com', headers= self.headers)
        except (requests.exceptions.ConnectionError) as e:
            logger.error(f'Function __init__ - Error - {e}')


    def pre_market_data(self, category):
        pre_market_category = {"NIFTY 50": "NIFTY", "Nifty Bank": "BANKNIFTY", "Emerge": "SME", "Securities in F&O":"FO", 
                            "Others": "OTHERS", "All": "ALL"}
        data = self.session.get(f'https://www.nseindia.com/api/market-data-pre-open?key={pre_market_category[category]}',
                                headers=self.headers).json()["data"]
        
        temp_data = []
        for i in data:
            temp_data.append(i["metadata"])
        df = pd.DataFrame(temp_data)
        df = df.set_index('symbol', drop=True)
        return df
    
    def equity_market_data(self, category, symbol_list=False):
        category = category.upper().replace(' ', '%20').replace('&', '%26')
        #logger.debug("Function equity_market_data")
        #data = self.session.get(f'https://www.nseindia.com/api/equity-stockindices?index={category}', headers=self.headers).json()["data"]
        try:
            #response = self.session.get(f'https://www.nseindia.com/api/equity-stockindices?index={category}', headers=self.headers, cookies=requests.cookies)
            response = self.session.get(f'https://www.nseindia.com/api/equity-stockindices?index={category}', headers=self.headers)
            response.raise_for_status()
        except (requests.HTTPError, requests.exceptions.ConnectionError, requests.exceptions.ChunkedEncodingError, 
                requests.ConnectionError, requests.exceptions.HTTPError) as e:
            logger.error(f'Function equity_market_data - Error - {e}')
            #logger.debug(f'HTTPStatus is - {HTTPStatus.description()}')
            #logger.debug(f'Exception Status code - {e.response.status_code()}')
            #self.__init__()
            return None
        if response.status_code != 401:
            try:
                #logger.debug(response.text.encode("utf-8"))
                logger.debug(response.encoding)
                #logger.debug(response.content)                
                #logger.debug(json.dumps(response.text))
                #logger.debug(json.loads(response.text))
                logger.debug(response.headers)
                #logger.debug(requests.head(f'https://www.nseindia.com/api/equity-stockindices?index={category}'))
                logger.debug(response.request)
                logger.debug(response.url)
                logger.debug(response.status_code)
                data = response.json()["data"]
                logger.debug(data)
                df = pd.DataFrame(data)
                logger.debug(df)
                df = df.drop("meta", axis=1)
                logger.debug(df)
                df = df.set_index('symbol', drop=True)
                logger.debug(df)
                if symbol_list:
                    logger.debug(f'if - symbol_list')
                    return list(df.index)
                else:
                    logger.debug(f'else - symbol_list')
                    return df
            except (requests.JSONDecodeError,requests.exceptions.JSONDecodeError,ValueError,KeyError) as e:
                logger.error(f'Function equity_market_data - Decoding JSON has failed - {e}')
                return None
        else:
            return None
        
    def about_holidays(self, category):
        data = self.session.get(f'https://www.nseindia.com/api/holiday-master?type={category.lower()}', headers=self.headers).json()
        df = pd.DataFrame(list(data.values())[0])
        return df
    
    def equity_info(self, symbol, trade_info=False):        
        symbol = symbol.replace(' ', '%20').replace('&', '%26')
        #data = self.session.get("https://www.nseindia.com/api/quote-equity?symbol=" + symbol + ("&section=trade_info" if trade_info else ""), headers=self.headers).json()
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
            except (ValueError,KeyError) as e:
                logger.error(f'Function equity_info - Decoding JSON has failed - {e}')
                return None
        else:
            return None            
    
    def futures_data(self, symbol, indices=False):
        symbol = symbol.replace(' ', '%20').replace('&', '%26')
        #data = self.session.get("https://www.nseindia.com/api/quote-derivative?symbol=" + symbol, headers=self.headers).json()
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
            except (ValueError,KeyError) as e:
                logger.error(f'Function futures_data - Decoding JSON has failed - {e}')
                return None
        else:
            return None

    def derivatives_data(self, symbol):
        symbol = symbol.replace(' ', '%20').replace('&', '%26')
        #data = self.session.get("https://www.nseindia.com/api/quote-derivative?symbol=" + symbol, headers=self.headers).json()
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
            except (ValueError,KeyError) as e:
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
        #data = self.session.get(url, headers=self.headers).json()["records"]
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
            except (ValueError,KeyError) as e:
                logger.error(f'Function options_data - Decoding JSON has failed - {e}')
                return None
        else:
            return None   