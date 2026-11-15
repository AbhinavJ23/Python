import pandas as pd
import numpy as np
from datetime import datetime
from baselogger import logger

def process_market_pass(df_live_ticks, sr_dataframe, previous_pass_data, config):
    """
    Phase 1 Upgraded Engine:
    1. Computes true interval deltas for Volume and Turnover (in Crores).
    2. Tracks rolling 5-interval average volume for 1.5x RVOL surge detection.
    3. Validates breakouts against S/R, Intraday VWAP, 1.5x RVOL, and Day High/Low proximity.
    """
    logger.debug("Processing market pass with live ticks and support/resistance baselines.")

    timestamp_str = datetime.now().strftime("%H:%M:%S")
    
    max_turnover_list = []
    min_turnover_list = []
    puvu_list = []
    pdvd_list = []
    puvdn_list = []
    pdvup_list = []
    
    current_pass_records = {}
    
    # Index Support/Resistance dataframe by Stock Symbol for fast lookups
    sr_indexed = None
    if sr_dataframe is not None and not sr_dataframe.empty and "Stock Symbol" in sr_dataframe.columns:
        sr_indexed = sr_dataframe.set_index("Stock Symbol")

    if df_live_ticks is None or df_live_ticks.empty or "Error" in df_live_ticks.columns:
        return {
            "max_turnover_df": pd.DataFrame(),
            "min_turnover_df": pd.DataFrame(),
            "puvu_list": [],
            "pdvd_list": [],
            "puvdn_list": [],
            "pdvup_list": [],
            "raw_pass_state": previous_pass_data
        }

    # Note: config['max_turnover_limit'] and config['min_turnover_limit'] are expected in raw Rupees.
    # If your sidebar stores them directly in Crores, remove the "/ 10000000" division on the limit variables.
    max_limit_cr = config.get("max_turnover_limit", 0) / 10000000
    min_limit_cr = config.get("min_turnover_limit", 0) / 10000000

    # Extract index change passed from app.py (defaults to 0.0 if missing)
    index_change_pct = float(config.get("index_change_pct", 0.0))
    rs_threshold = float(config.get("rs_threshold", 0.50))
    rvol_multiplier = float(config.get("rvol_multiplier", 1.50))
    hl_prox_pct = float(config.get("hl_proximity_pct", 0.25)) / 100.0  # Converts 0.25% -> 0.0025

    for _, row in df_live_ticks.iterrows():
        sym = row["Stock Symbol"]
        ltp = row["Last Price"]
        stock_change_pct = row["Change (%)"]
        vwap = row["VWAP"]
        day_high = row["Day High"]
        day_low = row["Day Low"]
        cum_vol = int(row["Volume"])
        cum_turnover = row["Turnover"]  # Total turnover cumulative since market open
        
        prev_metrics = previous_pass_data.get(sym, {})
        
        # Calculate Intraday Relative Strength (Alpha vs. Benchmark Index)
        rs_alpha = round(stock_change_pct - index_change_pct, 2)
        
        # --- PASS 1 BASELINE INITIALIZATION ---
        if not prev_metrics:
            current_pass_records[sym] = {
                "price": ltp,
                "cum_volume": cum_vol,
                "cum_turnover": cum_turnover,
                "interval_vol": 0,
                "interval_turnover": 0.0,
                "vol_history": []
            }
            continue

        prev_price = prev_metrics.get("price", ltp)
        prev_cum_vol = prev_metrics.get("cum_volume", 0)
        prev_cum_turnover = prev_metrics.get("cum_turnover", 0.0)
        prev_interval_vol = prev_metrics.get("interval_vol", 0)
        vol_history = list(prev_metrics.get("vol_history", []))

        # --- CALCULATE TRUE 5-MINUTE INTERVAL DELTAS ---
        interval_vol = cum_vol - prev_cum_vol
        if interval_vol < 0:
            interval_vol = cum_vol

        interval_turnover_raw = cum_turnover - prev_cum_turnover
        if interval_turnover_raw < 0:
            interval_turnover_raw = cum_turnover

        interval_turnover_cr = round(interval_turnover_raw / 10000000, 2)
        
        # Calculate interval price shift (%)
        alert_pct_threshold = float(config.get("price_alert_percentage", 0.0))

        # Calculate interval price shift (%)
        price_pct_shift = ((ltp - prev_price) / prev_price * 100) if prev_price > 0 else 0.0

        # Compare percentage shift against your sidebar alert threshold
        if price_pct_shift >= alert_pct_threshold and price_pct_shift > 0:
            price_status = "GREEN"
        elif price_pct_shift <= -alert_pct_threshold and price_pct_shift < 0:
            price_status = "RED"
        else:
            price_status = "NEUTRAL"

        # --- CALCULATE ROLLING AVERAGE VOLUME (RVOL BASELINE) ---
        avg_prev_vol = (sum(vol_history) / len(vol_history)) if len(vol_history) > 0 else prev_interval_vol

        # Update rolling window (keep last 5 completed intervals)
        if interval_vol > 0:
            vol_history.append(interval_vol)
            if len(vol_history) > 5:
                vol_history.pop(0)

        # Save current metrics and interval deltas for the NEXT pass comparison
        current_pass_records[sym] = {
            "price": ltp,
            "cum_volume": cum_vol,
            "cum_turnover": cum_turnover,
            "interval_vol": interval_vol,
            "interval_turnover": interval_turnover_raw,
            "vol_history": vol_history
        }

        # --- 1. MAX & MIN CUMULATIVE TURNOVER TABS (TABS 3 & 4) ---
        if interval_turnover_cr >= max_limit_cr and max_limit_cr > 0:
            logger.debug(f"Stock {sym} passed MAX turnover filter: Interval Turnover={interval_turnover_cr} Cr, Price Shift={price_pct_shift}%")
            max_turnover_list.append({
                "Time": timestamp_str,
                "Stock Name": sym,
                "Cumulative Turnover (Interval)": interval_turnover_cr,
                "Price": round(ltp, 2),
                "Price Shift (%)": round(price_pct_shift, 2),
                "Status": price_status
            })

        if 0 < interval_turnover_cr <= min_limit_cr:
            logger.debug(f"Stock {sym} passed MIN turnover filter: Interval Turnover={interval_turnover_cr} Cr, Price Shift={price_pct_shift}%")
            min_turnover_list.append({
                "Time": timestamp_str,
                "Stock Name": sym,
                "Cumulative Turnover (Interval)": interval_turnover_cr,
                "Price": round(ltp, 2),
                "Price Shift (%)": round(price_pct_shift, 2),
                "Status": price_status
            })

        # Skip interval-over-interval volume acceleration check on Pass 2 if Pass 1 baseline was 0
        if prev_interval_vol == 0 or avg_prev_vol == 0:
            continue

        # --- 2. SUPPORT / RESISTANCE BREAKOUT CHECK ---
        sr_signal = "YELLOW"
        if sr_indexed is not None and sym in sr_indexed.index:
            r1 = float(sr_indexed.loc[sym, "Resistance 1 (R1)"])
            s1 = float(sr_indexed.loc[sym, "Support 1 (S1)"])

            # Filter A: 1.5x Relative Volume Surge
            is_rvol_surge = interval_vol >= (rvol_multiplier * avg_prev_vol)
            
            # Filter B: Intraday VWAP Alignment (fallback to True if exchange VWAP is 0)
            is_above_vwap = (ltp > vwap) if vwap > 0 else True
            is_below_vwap = (ltp < vwap) if vwap > 0 else True
            
            # Filter C: High/Low Candle Proximity (within 0.25% of Day High / Day Low)
            is_near_high = (ltp >= day_high * (1.0 - hl_prox_pct)) if day_high > 0 else True
            is_near_low = (ltp <= day_low * (1.0 + hl_prox_pct)) if day_low > 0 else True

            # Filter D: Intraday Relative Strength (Alpha vs. Index)
            is_rs_outperforming = rs_alpha >= rs_threshold
            is_rs_underperforming = rs_alpha <= -rs_threshold

            # Combined Buy & Sell Signal Validation
            if ltp > r1 and is_rvol_surge and is_above_vwap and is_near_high and is_rs_outperforming:
                sr_signal = "GREEN"
            elif ltp < s1 and is_rvol_surge and is_below_vwap and is_near_low and is_rs_underperforming:
                sr_signal = "RED"

        # Unified dictionary containing Turnover (Cr) for all 4 signal tabs & Tab 10 alerts
        record_dict = {
            "Time": timestamp_str,
            "Stock Name": sym,
            "Price": round(ltp, 2),
            "Alpha (%)": rs_alpha,
            "Volume": interval_vol,
            "Turnover (Cr)": interval_turnover_cr,
            "Signal": sr_signal
        }

        # --- 3. FOUR DIVERGENCE VECTORS (TABS 6, 7, 8, 9) ---
        if ltp > prev_price and interval_vol > prev_interval_vol:
            puvu_list.append(record_dict)
            
        elif ltp < prev_price and interval_vol < prev_interval_vol:
            pdvd_list.append(record_dict)
            
        elif ltp > prev_price and interval_vol < prev_interval_vol:
            puvdn_list.append(record_dict)
            
        elif ltp < prev_price and interval_vol > prev_interval_vol:
            pdvup_list.append(record_dict)

    return {
        "max_turnover_df": pd.DataFrame(max_turnover_list),
        "min_turnover_df": pd.DataFrame(min_turnover_list),
        "puvu_list": puvu_list,
        "pdvd_list": pdvd_list,
        "puvdn_list": puvdn_list,
        "pdvup_list": pdvup_list,
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
