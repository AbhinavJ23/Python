import NseIndia
import pprint

nse = NseIndia.NSE()

#print(nse.pre_market_data("NIFTY 50"))
#print(nse.equity_market_categories)
print(nse.equity_market_data("NIFTY 50"))
#print(nse.holiday_categories)
#pprint.pprint(nse.equity_info('INFY', trade_info=True))``
#print(nse.futures_data('INFY', indices=False))
#print(nse.options_data("NIFTY", indices=True))