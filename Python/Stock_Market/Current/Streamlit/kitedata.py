import pandas as pd
import numpy as np
from datetime import datetime
from baselogger import logger

def process_market_pass(df_live_ticks, df_sr_baselines, previous_pass_data, config):
    
    logger.debug("Processing market pass with live ticks and support/resistance baselines.")

    """
    config dict keys expected: 
      - 'max_turnover_limit' (e.g., 100000000 for 10 Crore)
      - 'min_turnover_limit' (e.g., 20000000 for 2 Crore)
      - 'price_alert_percentage' (e.g., 0.5 for 0.5%)
    """
    timestamp_str = datetime.now().strftime("%H:%M:%S")
    
    # 1. Clean and transform incoming tick matrix data
    current_pass_records = {}
    max_turnover_list = []
    min_turnover_list = []
    
    # Merge support/resistance baselines for easy matrix operations
    sr_indexed = df_sr_baselines.set_index("Stock Symbol") if not df_sr_baselines.empty else None

    for _, row in df_live_ticks.iterrows():
        sym = row["Stock Symbol"]
        ltp = row["Last Price"]
        cum_volume = row["Volume"]
        cum_turnover = row["Turnover"]  # Total turnover cumulative since market open
        
        # Calculate Delta (Turnover/Volume since LAST 5-minute pass)
        prev_stock_data = previous_pass_data.get(sym, {})
        prev_cum_turnover = prev_stock_data.get("cum_turnover", 0)
        prev_cum_volume = prev_stock_data.get("cum_volume", 0)
        prev_price = prev_stock_data.get("price", ltp)
        
        # Interval turnover and volume calculation
        interval_turnover = cum_turnover - prev_cum_turnover
        interval_volume = cum_volume - prev_cum_volume
        
        # Fallback handling for first pass of the day
        if interval_turnover < 0: interval_turnover = cum_turnover
        if interval_volume < 0: interval_volume = cum_volume
            
        # Store for next loop cross-comparison sequence
        current_pass_records[sym] = {
            "price": ltp,
            "cum_turnover": cum_turnover,
            "cum_volume": cum_volume
        }
        
        # Calculate price percentage shift from previous pass
        price_pct_shift = ((ltp - prev_price) / prev_price) * 100 if prev_price > 0 else 0.0
        price_status = "GREEN" if price_pct_shift >= config['price_alert_percentage'] else "RED"
        
        # --- LOGIC 1: MAX CUMULATIVE TURNOVER FILTER ---
        if interval_turnover >= config['max_turnover_limit']:
            logger.debug(f"Stock {sym} passed MAX turnover filter: Interval Turnover={interval_turnover}, Price Shift={price_pct_shift}%")
            max_turnover_list.append({
                "Time": timestamp_str,
                "Stock Name": sym,
                "Cumulative Turnover (Interval)": round(interval_turnover/10000000, 2),
                "Price": round(ltp,2),
                "Price Shift (%)": round(price_pct_shift, 2),
                "Status": price_status
            })
            
        # --- LOGIC 2: MIN CUMULATIVE TURNOVER FILTER ---
        if interval_turnover <= config['min_turnover_limit'] and interval_turnover > 0:
            logger.debug(f"Stock {sym} passed MIN turnover filter: Interval Turnover={interval_turnover}, Price Shift={price_pct_shift}%")
            min_turnover_list.append({
                "Time": timestamp_str,
                "Stock Name": sym,
                "Cumulative Turnover (Interval)": round(interval_turnover/10000000, 2),
                "Price": round(ltp,2),
                "Price Shift (%)": round(price_pct_shift, 2),
                "Status": price_status
            })

    # Convert to standard DataFrames
    df_max_turnover = pd.DataFrame(max_turnover_list)
    df_min_turnover = pd.DataFrame(min_turnover_list)
    
    # --- LOGIC 4 & 5: PRICE UP/VOL UP & PRICE DOWN/VOL DOWN SIGNAL PARSER ---
    puvu_list = []
    pdvd_list = []
    
    for sym, current_metrics in current_pass_records.items():
        prev_metrics = previous_pass_data.get(sym, {})
        if not prev_metrics:
            continue  # Requires at least one previous pass to run comparisons
            
        # Extract previous pass parameters
        prev_price = prev_metrics.get("price", 0)
        prev_vol = prev_metrics.get("cum_volume", 0)
        
        curr_price = current_metrics["price"]
        curr_vol = current_metrics["cum_volume"]
        
        # Evaluate Support/Resistance alerts if index matches match
        sr_signal = "YELLOW"  # Default
        if sr_indexed is not None and sym in sr_indexed.index:
            r1 = sr_indexed.loc[sym, "Resistance 1 (R1)"]
            s1 = sr_indexed.loc[sym, "Support 1 (S1)"]
            if curr_price > r1:
                sr_signal = "GREEN"  # Breakout past resistance
            elif curr_price < s1:
                sr_signal = "RED"    # Breakdown below support

        # Condition 4: Price Up & Volume Up
        if curr_price > prev_price and curr_vol > prev_vol:
            logger.debug(f"Price Up & Volume Up detected for {sym}: Prev Price={prev_price}, Curr Price={curr_price}, Prev Vol={prev_vol}, Curr Vol={curr_vol}")
            puvu_list.append({"Time": timestamp_str, "Stock Name": sym, "Price": curr_price, "Volume": curr_vol, "Signal": sr_signal})
            
        # Condition 5: Price Down & Volume Down
        elif curr_price < prev_price and curr_vol < prev_vol:
            logger.debug(f"Price Down & Volume Down detected for {sym}: Prev Price={prev_price}, Curr Price={curr_price}, Prev Vol={prev_vol}, Curr Vol={curr_vol}")
            pdvd_list.append({"Time": timestamp_str, "Stock Name": sym, "Price": curr_price, "Volume": curr_vol, "Signal": sr_signal})

    return {
        "max_turnover_df": df_max_turnover,
        "min_turnover_df": df_min_turnover,
        "puvu_list": puvu_list,
        "pdvd_list": pdvd_list,
        "raw_pass_state": current_pass_records
    }

def generate_transposed_dashboard(historical_pass_list, symbol_list):
    """
    Constructs a transposed dashboard matrix:
    - Rows: Stock Symbols
    - Columns: Time Intervals (e.g., 09:33 P, 09:33 V) sorted chronologically
    """
    if not historical_pass_list:
        return pd.DataFrame()

    # 1. Initialize data rows with your active watchlist symbols
    matrix_data = {sym: {"Stock Symbol": sym} for sym in symbol_list}
    
    # 2. Loop through every completed historical 5-minute interval pass
    for pass_data in historical_pass_list:
        t_stamp = pass_data["timestamp"]  # e.g., "09:38"
        ticks = pass_data["ticks"]
        
        for sym in symbol_list:
            if sym in ticks:
                # Assign distinct sub-columns for Price (P) and Volume (V) per time block
                matrix_data[sym][f"{t_stamp} P"] = round(ticks[sym]["price"],2)
                matrix_data[sym][f"{t_stamp} V"] = ticks[sym]["cum_volume"]
            else:
                # Graceful fallback handles stocks that missed a tick
                matrix_data[sym][f"{t_stamp} P"] = 0.0
                matrix_data[sym][f"{t_stamp} V"] = 0

    # Convert our structured dictionary into a clean Pandas table frame
    df_transposed = pd.DataFrame(list(matrix_data.values()))
    
    # 3. Add a dynamic 'Total Vol' column by summing all volume columns
    volume_cols = [c for c in df_transposed.columns if c.endswith(" V")]
    if volume_cols:
        df_transposed["Total Volume Traded"] = df_transposed[volume_cols].sum(axis=1)

    logger.debug("Generated transposed dashboard matrix with historical data.")
    return df_transposed
