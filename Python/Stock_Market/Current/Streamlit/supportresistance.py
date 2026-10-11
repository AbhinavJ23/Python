import pandas as pd

def calculate_static_sr_levels(symbol_list, kite_instance):
    """
    Fetches real OHLC data for the stock list,
    calculates simple Pivot Point S/R lines, and returns a clean DataFrame.
    """
    sr_records = []
    
    # 1. Format symbols for API (e.g., 'RELIANCE' -> 'NSE:RELIANCE')
    formatted_symbols = [f"NSE:{sym}" if not sym.startswith("NSE:") else sym for sym in symbol_list]
    
    try:
        # 2. Bulk fetch yesterday's OHLC snapshot
        # kite.ohlc() accepts up to 200 instruments in a single batch call.
        ohlc_data = kite_instance.ohlc(formatted_symbols)
        
        for sym in symbol_list:
            api_key_name = f"NSE:{sym}"
            
            if api_key_name in ohlc_data:
                # Extract historical data attributes safely
                historical = ohlc_data[api_key_name]["ohlc"]
                high = historical["high"]
                low = historical["low"]
                close = historical["close"]
                
                # If market is currently open or previous data is 0, handle edge cases
                if high == 0 or low == 0:
                    continue
                
                # 3. Simple Pivot Point S/R Mathematical Formulas
                pivot = (high + low + close) / 3
                r1 = (2 * pivot) - low
                r2 = pivot + (high - low)
                s1 = (2 * pivot) - high
                s2 = pivot - (high - low)
                
                sr_records.append({
                    "Stock Symbol": sym,
                    "Prev Close": round(close, 2),
                    "Pivot Point (PP)": round(pivot, 2),
                    "Resistance 1 (R1)": round(r1, 2),
                    "Resistance 2 (R2)": round(r2, 2),
                    "Support 1 (S1)": round(s1, 2),
                    "Support 2 (S2)": round(s2, 2)
                })
                
    except Exception as api_error:
        # Fallback to a structured error state if API fails or connection drops
        return pd.DataFrame([{"Error": f"Failed to fetch data : {str(api_error)}"}])
        
    # Return as a clean structured Pandas DataFrame matrix
    return pd.DataFrame(sr_records)