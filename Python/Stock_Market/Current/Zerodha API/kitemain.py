from kitelogin import KiteLogin
from kiteconnect import KiteConnect
import pandas as pd
from baselogger import logger

class KiteMain:
    equity_market_categories = ['NIFTY 50', 'NIFTY NEXT 50', 'NIFTY 100', 'NIFTY 200', 'NIFTY 500', 'NIFTY MIDCAP 50', 'NIFTY MIDCAP 100', 'NIFTY SMALLCAP 100', 'NIFTY SMALLCAP 50', 'NIFTY SMALLCAP 250', 'NIFTY TOTAL MARKET', 'NIFTY BANK', 'NIFTY AUTO', 'NIFTY FMCG', 'NIFTY IT', 'NIFTY METAL', 'NIFTY PHARMA', 'NIFTY PSU BANK', 'NIFTY PRIVATE BANK', 'NIFTY REALTY', 'NIFTY HEALTHCARE INDEX', 'NIFTY CONSUMER DURABLES', 'NIFTY OIL & GAS']
    def __init__(self):
        self.kite_login = KiteLogin()
        self.api_key = self.kite_login.get_api_key()
        self.access_token = self.kite_login.get_access_token()
        self.kite = KiteConnect(api_key=self.api_key)
        self.kite.set_access_token(self.access_token)

    def get_nifty50_market_data(self):
        # Get all NSE instruments
        #instruments = self.kite.instruments(exchange="NSE")
        # Static NIFTY 50 symbols list
        nifty50_symbols = [
            "ADANIENT", "ADANIPORTS", "APOLLOHOSP", "ASIANPAINT", "AXISBANK", "BAJAJ-AUTO",
            "BAJFINANCE", "BAJAJFINSV", "BEL", "BHARTIARTL",
            "CIPLA", "COALINDIA", "DRREDDY", "EICHERMOT", "ETERNAL",
            "GRASIM", "HCLTECH", "HDFCBANK", "HDFCLIFE", "HEROMOTOCO",
            "HINDALCO", "HINDUNILVR", "ICICIBANK", "INDUSINDBK",
            "INFY", "ITC", "JIOFIN", "JSWSTEEL", "KOTAKBANK", "LT",
            "M&M", "MARUTI", "NESTLEIND", "NIFTY 50", "NTPC", "ONGC",
            "POWERGRID", "RELIANCE", "SBILIFE", "SBIN", "SHRIRAMFIN", "SUNPHARMA",
            "TATACONSUM", "TATAMOTORS", "TATASTEEL", "TCS", "TECHM",
            "TITAN", "TRENT", "ULTRACEMCO", "WIPRO"
        ]
        # Prepend 'NSE:' to each symbol
        nifty50_symbols = [f"NSE:{symbol}" for symbol in nifty50_symbols]
        #nifty = self.kite.quote('NSE:NIFTY 50')#['NSE:NIFTY 50']
        #print(nifty)
        market_data = []
        try:
            for sym in nifty50_symbols:
                market_data.append(self.kite.quote(sym)[sym])
        except Exception as e:
            logger.error(f"Error getting market data: {e}")
            return None

        # Prepare market data dictionary
        market_data_dict = {
            'symbol': [],
            'open': [],
            'high': [],
            'low': [],
            'close': [],
            'last_price': [],
            'change': [],
            'percent_change': [],
            'timestamp': [],
            'buy_quantity': [],
            'sell_quantity': [],
            'volume': [],
            'turnover': []
        }
        counter = 0
        for data in market_data:
            market_data_dict['symbol'].append(nifty50_symbols[counter][4:])  # Remove 'NSE:' prefix
            #market_data_dict['symbol'].append(market_data[counter].keys()) #.split(':')[1])
            market_data_dict['open'].append(data['ohlc']['open'])
            market_data_dict['high'].append(data['ohlc']['high'])
            market_data_dict['low'].append(data['ohlc']['low'])
            market_data_dict['close'].append(data['ohlc']['close'])
            market_data_dict['last_price'].append(data['last_price'])
            market_data_dict['change'].append(data['last_price'] - data['ohlc']['close'])
            market_data_dict['percent_change'].append(round(((data['last_price'] - data['ohlc']['close']) / data['ohlc']['close']) * 100, 2))
            market_data_dict['timestamp'].append(data.get('timestamp', None))
            if nifty50_symbols[counter] != 'NSE:NIFTY 50':
                market_data_dict['buy_quantity'].append(data['buy_quantity'])
                market_data_dict['sell_quantity'].append(data['sell_quantity'])
                market_data_dict['volume'].append(data['volume'])
                market_data_dict['turnover'].append(data['volume']*data['last_price'])
            else:
                market_data_dict['buy_quantity'].append(None)
                market_data_dict['sell_quantity'].append(None)
                market_data_dict['volume'].append(None)
                market_data_dict['turnover'].append(None)
            counter += 1

        market_data_df = pd.DataFrame(market_data_dict)
        market_data_df.set_index('symbol', inplace=True)
        #logger.debug(market_data_df)
        return market_data_df