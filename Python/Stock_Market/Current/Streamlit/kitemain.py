import pandas as pd
from datetime import datetime

def fetch_live_quotes(symbol_list, index_name, kite_instance):
    """
    Accepts watchlist list, active tracking index name, and your authorized kite connection object.
    Queries Zerodha for fresh multi-second quotes and returns a parsed DataFrame.
    """
    # 1. Format symbols for the Zerodha network (e.g., 'RELIANCE' -> 'NSE:RELIANCE')
    formatted_symbols = [f"NSE:{sym}" if not sym.startswith("NSE:") else sym for sym in symbol_list]

    index_token = f"NSE:{index_name.upper()}"
    if index_token not in formatted_symbols:
        formatted_symbols.append(index_token)
    
    try:
        # 2. Bulk pull live tick values (Prices, Volume, Open Interest, Turnover)
        raw_quotes = kite_instance.quote(formatted_symbols)
              
        parsed_records = []
        index_data = {"last_price": 0.0, "change": 0.0}

        for symbol_key, metrics in raw_quotes.items():
            # Check if this row is the master index metric block
            if symbol_key == index_token:
                index_data["last_price"] = round(metrics.get("last_price", 0.0), 2)
                index_data["change"] = round(metrics.get("change", 0.0), 2)
                continue
            # Strip out 'NSE:' prefix for a clean visual table layout
            clean_symbol = symbol_key.replace("NSE:", "")
            
            # Map parameters matching your exact screening requirements
            parsed_records.append({
                "Stock Symbol": clean_symbol,
                #"Time": metrics.get("timestamp", ""),
                "Time": datetime.strftime(metrics.get("timestamp", datetime.now()), "%d-%m-%Y %H:%M:%S") if metrics.get("timestamp") else "",
                "Open": round(metrics.get("ohlc", {}).get("open", 0.0), 2),
                "Day High": round(metrics.get("ohlc", {}).get("high", 0.0), 2),
                "Day Low": round(metrics.get("ohlc", {}).get("low", 0.0), 2),
                "Previous Close": round(metrics.get("ohlc", {}).get("close", 0.0), 2),
                "Last Price": round(metrics.get("last_price", 0.0), 2),
                "Change": round(metrics.get("last_price", 0.0) - metrics.get("ohlc", {}).get("close", 0.0), 2),
                "Change (%)": round((metrics.get("last_price", 0.0) - metrics.get("ohlc", {}).get("close", 0.0)) / metrics.get("ohlc", {}).get("close", 1.0) * 100, 2) if metrics.get("ohlc", {}).get("close", 0.0) != 0 else 0.0,
                "Volume": metrics.get("volume", 0),
                "Turnover": round(metrics.get("volume", 0) * metrics.get("last_price", 0.0), 2), # Cumulative turnover in INR
            })
            
        return pd.DataFrame(parsed_records), index_data
        
    except Exception as fetch_error:
        # Gracefully handle transient api drops without breaking the app engine
        return pd.DataFrame([{"Error": f"Live snapshot failure: {str(fetch_error)}"}])
