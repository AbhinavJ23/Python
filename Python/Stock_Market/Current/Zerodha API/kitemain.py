from kitelogin import KiteLogin
from kiteconnect import KiteConnect
import pandas as pd
from baselogger import logger
from nseindex import NseIndex

class KiteMain:
    #equity_market_categories = ['NIFTY 50', 'NIFTY NEXT 50', 'NIFTY 100', 'NIFTY 200', 'NIFTY 500', 'NIFTY MIDCAP 50', 'NIFTY MIDCAP 100', 'NIFTY SMALLCAP 100', 'NIFTY SMALLCAP 50', 'NIFTY SMALLCAP 250', 'NIFTY TOTAL MARKET', 'NIFTY BANK', 'NIFTY AUTO', 'NIFTY FMCG', 'NIFTY IT', 'NIFTY METAL', 'NIFTY PHARMA', 'NIFTY PSU BANK', 'NIFTY PRIVATE BANK', 'NIFTY REALTY', 'NIFTY HEALTHCARE INDEX', 'NIFTY CONSUMER DURABLES', 'NIFTY OIL & GAS']
    def __init__(self):
        self.kite_login = KiteLogin()
        self.kite_login.load_credentials()
        self.kite_login.load_access_token()
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

    def get_equity_market_data(self, index_symbol):
        if index_symbol not in self.equity_market_categories:
            logger.error(f"Invalid index symbol: {index_symbol}")
            return None
        
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
            logger.error(f"Market data for {index_symbol} is not implemented.")
            return None

        # Prepend 'NSE:' to each symbol
        symbols = [f"NSE:{symbol}" for symbol in symbols]
        market_data = []
        try:
            for sym in symbols:
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
            market_data_dict['symbol'].append(symbols[counter][4:])  # Remove 'NSE:' prefix
            market_data_dict['open'].append(data['ohlc']['open'])
            market_data_dict['high'].append(data['ohlc']['high'])
            market_data_dict['low'].append(data['ohlc']['low'])
            market_data_dict['close'].append(data['ohlc']['close'])
            market_data_dict['last_price'].append(data['last_price'])
            market_data_dict['change'].append(data['last_price'] - data['ohlc']['close'])
            market_data_dict['percent_change'].append(round(((data['last_price'] - data['ohlc']['close']) / data['ohlc']['close']) * 100, 2))
            market_data_dict['timestamp'].append(data.get('timestamp', None))

            if symbols[counter] not in [f"NSE:{sym}" for sym in self.equity_market_categories]:
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
        return market_data_df