import streamlit as st
import time
import os
import json
from datetime import datetime, timedelta
import basedb
import baselogin
import kitelogin
import trialcheck
import supportresistance
import kitemain
import kitedata
import pandas as pd
from baselogger import logger

st.set_page_config(layout="wide", page_title="Market Screener Framework")

# -----------------------------------------------------------------
# 1. DATABASE AUTOMATED SEEDING INITIALIZATION
# -----------------------------------------------------------------
if "db_initialized" not in st.session_state:
    basedb.setup_database()  # Instantly sets up SQLite table schema
    st.session_state.db_initialized = True

# -----------------------------------------------------------------
# 2. APPLICATION STATE PERSISTENCE MACHINE
# -----------------------------------------------------------------
# Initialize session states for the API sequence
if "api_setup_done" not in st.session_state:
    # Check if JSON file already exists from a previous run
    st.session_state.api_setup_done = kitelogin.load_api_credentials() is not None

if "api_authenticated" not in st.session_state:
    st.session_state.api_authenticated = False

if "user_logged_in" not in st.session_state:
    st.session_state.user_logged_in = False

# If API keys are setup but we aren't marked authenticated yet, check the file cache
if st.session_state.api_setup_done and not st.session_state.api_authenticated:
    SESSION_FILE = "session.json"
    
    # 1. Check if the token file exists from an earlier run
    if os.path.exists(SESSION_FILE):
        with open(SESSION_FILE, "r") as f:
            session_data = json.load(f)
        
        # 2. Check if the cached session token was generated TODAY
        today_str = datetime.now().strftime("%Y-%m-%d")
        if session_data.get("date") == today_str:
            
            try:
                from kiteconnect import KiteConnect
                
                # Load your master credentials to establish the API link
                with open("credentials.json", "r") as credentials_file:
                    creds = json.load(credentials_file)
                    
                # 3. Re-initialize the active connection object using the cached token
                kite = KiteConnect(api_key=creds['api_key'])
                kite.set_access_token(session_data.get("access_token"))
                
                # 4. Save the live connection straight into Streamlit's active memory
                st.session_state.kite = kite
                st.session_state.api_authenticated = True
                
                # Pop up a clean, non-intrusive alert on the web browser screen
                st.toast("🔄 Reused valid daily Zerodha session token seamlessly!", icon="✅")
                
            except Exception as e:
                # If the file is corrupt or token expired early, delete it to force a clean login
                os.remove(SESSION_FILE)
                st.warning("Daily session token expired or corrupt. Please re-authenticate.")
# =========================================================================

# -----------------------------------------------------------------
# STAGE A: ONE-TIME ZERODHA API KEY & SECRET SETUP
# -----------------------------------------------------------------
if not st.session_state.api_setup_done:
    logger.debug("API keys not found. Prompting user for initial setup.")
    st.subheader("🔑 Step 1: Initial API Configuration")
    st.info("This is a one-time setup. Your keys will be saved locally in a JSON file.")
    
    api_key_input = st.text_input("Enter your API Key:")
    api_secret_input = st.text_input("Enter your API Secret:", type="password")
    
    if st.button("Save API Keys"):
        if api_key_input and api_secret_input:
            kitelogin.save_api_credentials(api_key_input, api_secret_input)
            st.session_state.api_setup_done = True
            logger.debug("API keys saved successfully. Proceeding to authentication stage.")
            st.success("API Keys saved successfully! Reloading...")
            time.sleep(1)
            st.rerun()
        else:
            st.error("Please fill in both fields before saving.")
    st.stop()  # Halts execution here until keys are stored

# -----------------------------------------------------------------
# STAGE B: DAILY ZERODHA API TOKEN AUTHENTICATION
# -----------------------------------------------------------------
elif not st.session_state.api_authenticated:
    logger.debug("API keys found. Prompting user for daily authentication.")
    st.subheader("🔐 Step 2: Daily API Authentication")
    
    # Load credentials from the JSON file
    creds = kitelogin.load_api_credentials()
    api_key = creds["api_key"]
    api_secret = creds["api_secret"]
    
    # Generate the web login link dynamically
    login_url = kitelogin.get_zerodha_login_url(api_key)
    
    st.markdown(f"### [👉 Click here to log in to and generate a Request Token]({login_url})")
    st.caption("Once you log in, copy the 'request_token' parameter out of the resulting browser address bar URL.")
    
    request_token = st.text_input("Paste your generated Request Token here:")
    
    if st.button("Authenticate  Session"):
        if request_token:
            try:
                # Exchange request token for access token using your function
                access_token, kite_instance = kitelogin.generate_session_token(
                    request_token, api_key, api_secret
                )
                logger.debug("Session token generated successfully.")
                with open("session.json", "w") as f:
                    json.dump({
                        "access_token": access_token,
                        "date": datetime.now().strftime("%Y-%m-%d")
                    }, f)
                # Store the authorized kite session instance safely inside memory
                st.session_state.kite = kite_instance
                st.session_state.api_authenticated = True
                logger.debug("Session token saved to session.json and stored in session state.")               
                st.success("Session Authenticated Successfully!")
                time.sleep(1)
                st.rerun()
            except Exception as e:
                st.error(f"Authentication Failed: {e}. Please double-check your token.")
        else:
            st.error("Please paste the request token first.")
    st.stop()

# -----------------------------------------------------------------
# SUCCESS: READY FOR USER LOGIN AND MAIN LOGIC
# -----------------------------------------------------------------
# -----------------------------------------------------------------
# STAGE C: USER ACCOUNT PROFILE ACCESS LAYERS (Replaces Tkinter Boxes)
# -----------------------------------------------------------------
elif not st.session_state.user_logged_in:
    st.subheader("👤 Step 3: Application Profile Identity and Index Selection")
    
    auth_action = st.radio("Access Operation Mode", ["Sign In", "Create New User"], horizontal=True)
    
    # 1. PROFILE SIGN-IN PORTAL
    if auth_action == "Sign In":
        with st.form("signin_form"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")

            import nseindex
            # Get the user-friendly names from our dictionary keys
            available_indices = list(nseindex.INDEX_MAPPING.keys())
            
            selected_index_name = st.selectbox(
                "🎯 Select Market Index to Track for this Session:",
                options=available_indices,
                index=0
            )

            login_clicked = st.form_submit_button("Log In")
            
            if login_clicked:
                if baselogin.valid_user(username):
                    if username and password:
                        if baselogin.verify_user_login(username, password):
                            st.session_state.user_logged_in = True
                            st.session_state.current_user = username
                            st.success(f"Access granted for user : {username}")
                            logger.debug(f"User '{username}' logged in successfully. Selected index: {selected_index_name}")
                            basedb.log_user_login(username)  # Log the successful login event
                            st.session_state.selected_index_name = selected_index_name
                            st.session_state.selected_symbols = nseindex.INDEX_MAPPING[selected_index_name]
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.error("❌ Invalid Username or Password.")
                    else:
                        st.error("Fields cannot be left blank.")
                else:
                    st.error("❌ Invalid Username. Please register with a valid username from the predefined list.")
                    
    # 2. LOCAL ACCOUNT CREATION PORTAL
    elif auth_action == "Create New User":
        logger.debug("User selected 'Create New User'. Displaying registration form.")
        #st.info("Creating a profile stores your secure hashed login data inside the local SQLite table layer.")
        with st.form("signup_form"):
            new_user = st.text_input("Choose Username")
            new_pass = st.text_input("Choose Password", type="password")
            conf_pass = st.text_input("Confirm Password", type="password")
            signup_clicked = st.form_submit_button("Register New User")
                        
            if signup_clicked:
                if new_user and new_pass and conf_pass:
                    if new_pass != conf_pass:
                        st.error("❌ Passwords do not match.")
                    else:
                        try:
                            # Pass web form variables cleanly to your database writing logic
                            if baselogin.valid_user(new_user):
                                baselogin.do_register(new_user, new_pass)
                                logger.debug(f"User '{new_user}' registered successfully.")
                                st.success("🎉 Registration successful! Toggle your selection option above to 'Sign In'.")
                            else:
                                st.error("❌ Invalid Username. Please choose a valid username from the predefined list to register.")
                        except Exception as reg_err:
                            st.error(f"❌ Failed to register user: {reg_err}")
                else:
                    st.error("All fields must be filled properly.")
    st.stop()

# -----------------------------------------------------------------
# STAGE D: THE MAIN MULTI-TAB WORKSPACE PREVIEW
# -----------------------------------------------------------------
else:
    # =========================================================================
    # 🔒 TRIAL VALIDATION SCREEN EXCLUSION GATE
    # =========================================================================
    is_trial_valid, trial_message, days_remaining = trialcheck.verify_trial_status()
    
    if not is_trial_valid:
        logger.warning("Trial period has expired. Terminating session and prompting user to log out.")
        st.error(trial_message)
        st.warning("🔒 This software package execution handle has been terminated.")
        
        if st.button("Log Out of Profile"):
            st.session_state.user_logged_in = False
            st.rerun()
        st.stop() 
    # =========================================================================

    with st.sidebar:
        st.success(f"👤 Profile: {st.session_state.current_user}")
        # DYNAMIC DAYS LEFT DISPLAY ---
        if days_remaining <= 5:
            # Highlight in warning orange/yellow if 5 days or fewer remain
            st.warning(f"⚠️ License Expires in: **{days_remaining} Days**")
        else:
            # Standard clean confirmation badge
            st.success(f"🔒 Trial Active: **{days_remaining} Days Left**")
        # -----------------------------------------------
        st.info(f"📊 Tracking Index: {st.session_state.selected_index_name}")
        #st.caption(f"Loaded Watchlist Count: {len(st.session_state.selected_symbols)} Stocks")

        #st.markdown("---")
        st.markdown("### ⚙️ Screening Configurations")
        
        # Configurable settings inputs replacing your static hardcoded variables
        duration_minutes = st.number_input("Pass Duration (Mins):", min_value=1, max_value=60, value=5)
        max_crore = st.number_input("Max Turnover Threshold (In Crores):", min_value=0.1, value=10.0, step=0.5)
        min_crore = st.number_input("Min Turnover Threshold (In Crores):", min_value=0.1, value=2.0, step=0.5)
        alert_pct = st.number_input("Price Comparison Target Alert (%)", min_value=0.01, value=0.5, step=0.05)
        
        # Convert Crores to absolute integer values for pure math comparison mapping
        runtime_config = {
            'max_turnover_limit': max_crore * 10000000,
            'min_turnover_limit': min_crore * 10000000,
            'price_alert_percentage': alert_pct
        }
        
        if st.button("🚪Log Out"):
            # Clear cached items to ensure full resets on profile swap
            if "sr_matrix_dataframe" in st.session_state:
                del st.session_state.sr_matrix_dataframe
            st.session_state.user_logged_in = False
            st.rerun()
    
    #INITIALIZE IN-MEMORY TIME-SERIES DATABASES
    if "sr_matrix_dataframe" not in st.session_state:
        logger.debug("Initializing in-memory time-series databases.")
        st.session_state.sr_matrix_dataframe = supportresistance.calculate_static_sr_levels(st.session_state.selected_symbols, st.session_state.kite)
        st.session_state.previous_pass_cache = {}
        st.session_state.dashboard_history = pd.DataFrame()
        st.session_state.puvu_history = pd.DataFrame()
        st.session_state.pdvd_history = pd.DataFrame()
        st.session_state.puvd_history = pd.DataFrame()
        st.session_state.pdvu_history = pd.DataFrame()
        #st.session_state.max_turnover_current = pd.DataFrame()
        #st.session_state.min_turnover_current = pd.DataFrame()
        st.session_state.max_turnover_history = pd.DataFrame() 
        st.session_state.min_turnover_history = pd.DataFrame()
        st.session_state.last_pass_timestamp = datetime.now()
        st.session_state.historical_passes = []

    #st.title("🚀 Real-Time Market Core Screener Matrix Workspace")
    #with st.container(border=False):
        # Combining text elements inside a single triple-quoted block strips all vertical gap blocks
        #st.markdown("#### 📈 Real-Time Market Data - Price, Volume, Turnover Monitoring")
    #st.markdown(f"#### 📈 Real-Time Market Data - Price,Volume,Turnover Monitoring")
    
    with st.container():
    # This targeted inline style ONLY strips the empty space above this specific header element
        st.markdown("""
        <div style='margin-top: -35px; padding-top: 0px; margin-bottom: 5px;'>
            <h4 style='font-size: 1.3rem !important; margin: 0px; padding: 0px;'>
                📈 Real-Time Market Data - Price, Volume, Turnover Monitoring
            </h4>
        </div>
        <hr style='margin-top: 2px; margin-bottom: 10px; border: 0; border-top: 1px solid #ccc;'>
    """, unsafe_allow_html=True)

    # --- CALCULATE REAL S/R MATRIX ONCE ON ENTERING HOME DASHBOARD ---
    if "sr_matrix_dataframe" not in st.session_state:
        with st.spinner("⏳ Fetching previous day candle limits and building S/R matrix..."):
            # Pass BOTH the selected stock arrays and the live initialized Zerodha Kite instance
            df_sr = supportresistance.calculate_static_sr_levels(
                st.session_state.selected_symbols, 
                st.session_state.kite
            )
            st.session_state.sr_matrix_dataframe = df_sr
            st.toast("🎯 Support & Resistance levels mapped successfully!", icon="✅")
            logger.debug("S/R matrix initialized successfully.")

    # NEW COMPACT SUMMARY LAYER: TOP 5 GAINERS & LOSERS OF THE DAY
    # SPLIT HEADER FOR LEADERS, INDEX AND LIVE CLOCK
    
    #col_title, col_clock = st.columns([3, 1])
    col_title, col_index, col_clock = st.columns([2, 2, 1])
    with col_title:
        st.markdown("##### 📊 Today's Market Leaders")
    
    with col_index:
        index_placeholder = st.empty()

    with col_clock:
        import streamlit.components.v1 as components
        
        components.html("""
            <div style="text-align: right; font-family: monospace; font-weight: bold; font-size: 1.15rem; color: #1e293b; margin-top: 5px;">
                🕒 <span id="live_clock">00:00:00 AM</span>
            </div>
            <script>
                function updateClock() {
                    const now = new Date();
                    let hours = now.getHours();
                    const minutes = String(now.getMinutes()).padStart(2, '0');
                    const seconds = String(now.getSeconds()).padStart(2, '0');
                    const ampm = hours >= 12 ? 'PM' : 'AM';
                    
                    hours = hours % 12;
                    hours = hours ? hours : 12; // Handle midnight '0' as '12'
                    const hoursStr = String(hours).padStart(2, '0');
                    
                    document.getElementById('live_clock').innerText = hoursStr + ':' + minutes + ':' + seconds + ' ' + ampm;
                }
                // Run instantly on load, then trigger updates exactly every 1000 milliseconds (1 second)
                updateClock();
                setInterval(updateClock, 1000);
            </script>
        """, height=35)

    # Establish a side-by-side split container window layout
    #col_gainers, col_losers = st.columns(2)
    col_gainers, col_losers, col_chart = st.columns([1, 1, 2])
    
    with col_gainers:
        st.markdown("<span style='color:#155724; font-weight:bold;'>▲ Top 5 Gainers</span>", unsafe_allow_html=True)
        gainers_placeholder = st.empty() # Allocation space frame for gainers matrix
        
    with col_losers:
        st.markdown("<span style='color:#721c24; font-weight:bold;'>▼ Top 5 Losers</span>", unsafe_allow_html=True)
        losers_placeholder = st.empty() # Allocation space frame for losers matrix

    with col_chart:
        st.markdown(f"<span style='color:#1e293b; font-weight:bold;'>📈 Live {st.session_state.selected_index_name} Intraday Trend</span>", unsafe_allow_html=True)
        chart_placeholder = st.empty() # Visual canvas allocation for the index trend line

    #st.markdown("---") # Visual separation line separating summaries from the 7 tabs

    # Creating the 4 Tab Views
    t1, t2, t3, t4, t5, t6, t7, t8, t9 = st.tabs([
        "🎯 Support / Resistance",
        "🔴 Live Market Data", 
        "🚀 Max Cum. Turnover",
        "📉 Min Cum. Turnover",
        "📊 Price Volume Dashboard",
        "📈 Price Up Volume Up",
        "📉 Price Down Volume Down",
        "⚡ Price Up Volume Down",
        "💥 Price Down Volume Up" 
    ])

    # DISPLAY THE DYNAMIC S/R CALCULATIONS FROM THE LIVE API DATA
    with t1:
        logger.debug("Rendering Support/Resistance matrix for selected index.")
        st.markdown(f"##### 📊 Watchlist Pivot Point Matrix: `{st.session_state.selected_index_name}`")
        #st.subheader(f"📊 Live Pivot Point Matrix: {st.session_state.selected_index_name}")
        #st.caption("Calculations generated using the previous trading day's true OHLC distribution.")
        
        # Verify if an error column exists before attempting to display tables
        if not st.session_state.sr_matrix_dataframe.empty and "Error" in st.session_state.sr_matrix_dataframe.columns:
            st.error(st.session_state.sr_matrix_dataframe.iloc[0]["Error"])
        else:
            st.dataframe(
                st.session_state.sr_matrix_dataframe, 
                use_container_width=True, 
                hide_index=True
            )    
    with t2:
        logger.debug("Rendering Live Market Data tab for selected index.")
        #st.subheader("Real-Time Multi-Second Ticker Monitoring")
        st.markdown(f"##### ⚡ Real-Time Multi-Second Stock Price Monitoring: `{st.session_state.selected_index_name}`")
        #st.write(f"Configured to monitor the **{st.session_state.selected_index_name}** universe.")
        live_ticker_frame = st.empty() 

    with t3:
        logger.debug("Rendering Max Cumulative Turnover tab for selected index.")
        st.markdown("##### 🚀 Stocks above Max Cumulative Turnover Limit")
        max_turnover_placeholder = st.empty()

    with t4:
        logger.debug("Rendering Min Cumulative Turnover tab for selected index.")
        st.markdown("##### 📉 Stocks below Min Cumulative Turnover Limit")
        min_turnover_placeholder = st.empty()

    with t5:
        logger.debug("Rendering Time-Series Price Volume Summary Dashboard tab for selected index.")
        st.markdown("##### 📊 Time-Series Price Volume Summary Dashboard")
        dashboard_history_placeholder = st.empty() 

    with t6:
        logger.debug("Rendering Cumulative Price & Volume Gaining Breakouts tab for selected index.")
        st.markdown("##### 📈 Cumulative Price & Volume Gaining Breakouts")
        puvu_placeholder = st.empty()

    with t7:
        logger.debug("Rendering Cumulative Price & Volume Dropping Breakdowns tab for selected index.")
        st.markdown("##### 📉 Cumulative Price & Volume Dropping Breakdowns")
        pdvd_placeholder = st.empty()

    with t8:
        logger.debug("Rendering Cumulative Price Up & Volume Down Divergences tab for selected index.")
        st.markdown("#### ⚡ Cumulative Price Up & Volume Down Divergences Log")
        puvd_placeholder = st.empty()

    with t9:
        logger.debug("Rendering Cumulative Price Down & Volume Up Divergences tab for selected index.")
        st.markdown("#### 💥 Cumulative Price Down & Volume Up Divergences Log")
        pdvu_placeholder = st.empty()
  
    if "index_chart_history" not in st.session_state:
        st.session_state.index_chart_history = []

    # -----------------------------------------------------------------
    # ⚡ AUTOMATED MULTI-SECOND POLLING AND LIVE INJECTION ENGINE
    # -----------------------------------------------------------------
    # This loop takes over and streams data continuously onto the screen
    while True:
        # A. Pull fresh ticks via kitemain module passing the active authorized workspace session
        df_live_ticks, index_metrics = kitemain.fetch_live_quotes(
            st.session_state.selected_symbols, 
            st.session_state.selected_index_name,
            st.session_state.kite
        )
        
        # Check if error occurred during fetch sequence
        if not df_live_ticks.empty and "Error" in df_live_ticks.columns:
            st.sidebar.error("⚠️ Connection error encountered.")
            time.sleep(2)
            continue
        
        logger.debug("Live ticks fetched successfully.")
        # Overwrite Tab 1 (Live Tickers Grid) instantly every 3 seconds
        live_ticker_frame.dataframe(df_live_ticks, use_container_width=True, hide_index=True)

        # -----------------------------------------------------------------
        # 👉 NEW: FORMAT AND INJECT LIVE INDEX SNAPSHOT WITH CONDITIONAL COLORING
        # -----------------------------------------------------------------
        idx_price = index_metrics["last_price"]
        idx_change = index_metrics["change"]
        
        # Determine green text for gaining or red for dropping index trends
        color_hex = "#155724" if idx_change >= 0 else "#721c24"
        arrow_sign = "▲" if idx_change >= 0 else "▼"
        plus_sign = "+" if idx_change >= 0 else ""
        
        # Overwrite the middle placeholder handle
        index_placeholder.markdown(
            f"""
            <div style='text-align: center; margin-top: 5px; font-family: sans-serif;'>
                <span style='font-weight: bold; color: #475569; font-size: 0.95rem;'>{st.session_state.selected_index_name}:</span>
                <span style='font-weight: 700; color: #1e293b; font-size: 1.15rem; margin-left: 5px;'>{idx_price:,.2f}</span>
                <span style='font-weight: bold; color: {color_hex}; font-size: 1.0rem; margin-left: 8px;'>
                    {arrow_sign} {plus_sign}{idx_change:.2f}%
                </span>
            </div>
            """, 
            unsafe_allow_html=True
        )

        # -----------------------------------------------------------------
        # CALCULATE AND INJECT TOP 5 MOVERS
        # -----------------------------------------------------------------
        if not df_live_ticks.empty and "Change (%)" in df_live_ticks.columns:
            # Drop unnecessary operational columns to optimize layout screen room
            display_cols = ["Stock Symbol", "Last Price", "Change (%)", "Volume"]
            df_movers = df_live_ticks[display_cols].copy()
            
            # 1. Compile Top 5 Gainers (Sort descending: highest positive returns first)
            df_top_gainers = df_movers.sort_values(by="Change (%)", ascending=False).head(5)
            
            # 2. Compile Top 5 Losers (Sort ascending: biggest negative returns first)
            df_top_losers = df_movers.sort_values(by="Change (%)", ascending=True).head(5)
            
            # Helper design to apply soft background green/red styling across the summary matrix rows
            style_g = "background-color: #d4edda; color: #155724; font-weight: bold;"
            style_l = "background-color: #f8d7da; color: #721c24; font-weight: bold;"
            
            styled_gainers = df_top_gainers.style.map(lambda v: style_g, subset=["Change (%)"]).format(precision=2)
            styled_losers = df_top_losers.style.map(lambda v: style_l, subset=["Change (%)"]).format(precision=2)
            
            # Update the placeholders safely out of the continuous data flow stream
            gainers_placeholder.dataframe(styled_gainers, use_container_width=True, hide_index=True)
            losers_placeholder.dataframe(styled_losers, use_container_width=True, hide_index=True)

            # -----------------------------------------------------------------
            # 📊 LIVE INDEX TIME-SERIES CHART GENERATOR
            # -----------------------------------------------------------------
            # Append the current index price along with the system time ticker
            st.session_state.index_chart_history.append({
                "Time": datetime.now().strftime("%H:%M:%S"),
                "Price": index_metrics["last_price"]
            })
            
            # Convert history stack to a temporary data frame mapping for rendering
            df_chart_data = pd.DataFrame(st.session_state.index_chart_history)
            
            # Prevent excessive memory consumption by keeping only the last 300 data points (~15 mins of data)
            if len(df_chart_data) > 300:
                st.session_state.index_chart_history.pop(0)
                
            # Render a native Streamlit line chart that fits perfectly in the right half
            if not df_chart_data.empty:
                chart_placeholder.line_chart(
                    df_chart_data.set_index("Time")["Price"],
                    use_container_width=True,
                    height=130 # Tight height parameter prevents displacing the 9 tabs below
                )

        # B. Check if time duration threshold has crossed to lock the current pass
        now = datetime.now()
        if now - st.session_state.last_pass_timestamp >= timedelta(minutes=duration_minutes):
            logger.debug("Duration threshold crossed. Executing market pass calculations.")
            # Execute calculation matrix using kitedata file engine
            pass_results = kitedata.process_market_pass(
                df_live_ticks, 
                st.session_state.sr_matrix_dataframe,
                st.session_state.previous_pass_cache,
                runtime_config
            )

            # Appends the completed pass state directly to the historical array log
            st.session_state.historical_passes.append({
                "timestamp": now.strftime("%H:%M:%S"),
                "ticks": pass_results["raw_pass_state"]
            })
            
            # C. Update dynamic state stores securely
            st.session_state.previous_pass_cache = pass_results["raw_pass_state"]
            #st.session_state.max_turnover_current = pass_results["max_turnover_df"]
            #st.session_state.min_turnover_current = pass_results["min_turnover_df"]

            if not pass_results["max_turnover_df"].empty:
                st.session_state.max_turnover_history = pd.concat([st.session_state.max_turnover_history, pass_results["max_turnover_df"]]).drop_duplicates().reset_index(drop=True)
                
            if not pass_results["min_turnover_df"].empty:
                st.session_state.min_turnover_history = pd.concat([st.session_state.min_turnover_history, pass_results["min_turnover_df"]]).drop_duplicates().reset_index(drop=True)
            
            # Append historical rows for break registers seamlessly
            if pass_results["puvu_list"]:
                logger.debug(f"Appending {len(pass_results['puvu_list'])} new Price Up Volume Up entries to history.")
                df_new_puvu = pd.DataFrame(pass_results["puvu_list"])
                st.session_state.puvu_history = pd.concat([st.session_state.puvu_history, df_new_puvu]).drop_duplicates().reset_index(drop=True)
                
            if pass_results["pdvd_list"]:
                logger.debug(f"Appending {len(pass_results['pdvd_list'])} new Price Down Volume Down entries to history.")
                df_new_pdvd = pd.DataFrame(pass_results["pdvd_list"])
                st.session_state.pdvd_history = pd.concat([st.session_state.pdvd_history, df_new_pdvd]).drop_duplicates().reset_index(drop=True)

            if pass_results["puvd_list"]:
                logger.debug(f"Appending {len(pass_results['puvd_list'])} new Price Up Volume Down entries to history.")
                st.session_state.puvd_history = pd.concat([st.session_state.puvd_history, pd.DataFrame(pass_results["puvd_list"])]).drop_duplicates().reset_index(drop=True)

            if pass_results["pdvu_list"]:
                logger.debug(f"Appending {len(pass_results['pdvu_list'])} new Price Down Volume Up entries to history.")
                st.session_state.pdvu_history = pd.concat([st.session_state.pdvu_history, pd.DataFrame(pass_results["pdvu_list"])]).drop_duplicates().reset_index(drop=True)
                
            # Reset pass tracking timer baseline checkpoint
            st.session_state.last_pass_timestamp = now
            logger.debug("Pass complete! Operational metrics refreshed across tabs.")
            for list_name, label in [
                ("puvu_list", "📈 Price Up Volume Up"),
                ("pdvd_list", "📉 Price Down Volume Down"),
                ("puvd_list", "⚡ Price Up Volume Down"),
                ("pdvu_list", "💥 Price Down Volume Up")
                ]:
                if pass_results.get(list_name):
                    for record in pass_results[list_name]:
                        # Only fire a popup if it breaches Support (RED) or Resistance (GREEN)
                        if record.get("Signal") in ["GREEN", "RED"]:
                            sym_name = record["Stock Name"]
                            ltp_val = record["Price"]
                            vol_val = record["Volume"]
                            #turn_val = record["Turnover (Cr)"]
                            alert_type = "🚀 RESISTANCE BREAKOUT" if record["Signal"] == "GREEN" else "⚠️ SUPPORT BREAKDOWN"
                                
                            # Fire an institutional-grade toast notification popup box instantly
                            st.toast(
                                f"**{alert_type}**\n\n"
                                f"📦 **Stock:** {sym_name}\n"
                                f"💵 **Price:** ₹{ltp_val:,.2f}\n"
                                f"📊 **Vol:** {vol_val:,}\n"
                                #f"💰 **Turnover:** {turn_val:.2f} Cr\n"
                                f"🔍 *Vector:* {label}",
                                icon="🎯"
                            )  
            #st.toast("⏳ Pass complete! Operational metrics refreshed across tabs.", icon="📊")

        # D. Display updated states onto respective visual tab frames with color profiles
        # Helper cell color highlighting function for turnover matrix grids
        def color_price_status(val):
            logger.debug(f"Coloring cell with value: {val}")
            if val == "GREEN": return "background-color: #d4edda; color: #155724; font-weight: bold;"
            if val == "RED": return "background-color: #f8d7da; color: #721c24; font-weight: bold;"
            return ""

        def color_sr_signals(row):
            logger.debug(f"Coloring S/R signal row: {row.to_dict()}")
            style = [''] * len(row)
            if row.get("Signal") == "GREEN": style = ['background-color: #d4edda; color: #155724; font-weight: bold;'] * len(row)
            elif row.get("Signal") == "RED": style = ['background-color: #f8d7da; color: #721c24; font-weight: bold;'] * len(row)
            elif row.get("Signal") == "YELLOW": style = ['background-color: #fff3cd; color: #856404;'] * len(row)
            return style

        # Render Tab 3
        #if not st.session_state.max_turnover_current.empty:
        #    logger.debug("Rendering Max Cumulative Turnover tab.")
        #    max_turnover_placeholder.dataframe(st.session_state.max_turnover_current.style.map(color_price_status, subset=["Status"]).format(precision=2), use_container_width=True, hide_index=True)
        #else:
        #    max_turnover_placeholder.info("No stocks have passed the Max Turnover threshold in this pass yet.")

        # Render Tab 4
        #if not st.session_state.min_turnover_current.empty:
        #    logger.debug("Rendering Min Cumulative Turnover tab.")
        #    min_turnover_placeholder.dataframe(st.session_state.min_turnover_current.style.map(color_price_status, subset=["Status"]).format(precision=2), use_container_width=True, hide_index=True)
        #else:
        #    min_turnover_placeholder.info("No stocks are below the Min Turnover threshold in this pass yet.")

        # --- RENDER TAB 3 (MAX TURNOVER GRID TIMEWISE HISTORICAL) ---
        if not st.session_state.max_turnover_history.empty:
            styled_max = st.session_state.max_turnover_history.style.map(
                color_price_status, subset=["Status"]
            ).format(precision=2)
            max_turnover_placeholder.dataframe(styled_max, use_container_width=True, hide_index=True)
        else:
            max_turnover_placeholder.info("Awaiting the first completed pass to register timewise Max Turnover records.")

        # --- RENDER TAB 4 (MIN TURNOVER GRID TIMEWISE HISTORICAL) ---
        if not st.session_state.min_turnover_history.empty:
            styled_min = st.session_state.min_turnover_history.style.map(
                color_price_status, subset=["Status"]
            ).format(precision=2)
            min_turnover_placeholder.dataframe(styled_min, use_container_width=True, hide_index=True)
        else:
            min_turnover_placeholder.info("Awaiting the first completed pass to register timewise Min Turnover records.")

        # Render Tab 5 (Excel Dashboard History Grid View)
        # -------------------------------------------------------------
        # D. APPLY DYNAMIC STRIP COLOR DESIGN HIGHLIGHTS TO DATAFRAMES
        # -------------------------------------------------------------
        
        if "historical_passes" in st.session_state and st.session_state.historical_passes:
            df_transposed_matrix = kitedata.generate_transposed_dashboard(
                st.session_state.historical_passes,
                st.session_state.selected_symbols
            )
            logger.debug("Rendering Price Volume Dashboard tab with transposed historical data.")
            def style_transposed_cells(df):
                logger.debug("Applying dynamic strip color design highlights to transposed cells.")
                styles = pd.DataFrame('', index=df.index, columns=df.columns)
                price_cols = sorted([c for c in df.columns if c.endswith(" P")])
                for idx, col in enumerate(price_cols):
                    if idx == 0: continue
                    prior_col = price_cols[idx - 1]
                    corresponding_vol_col = col.replace(" P", " V")
                    for row_idx in df.index:
                        curr_p = df.at[row_idx, col]
                        prev_p = df.at[row_idx, prior_col]
                        if curr_p > prev_p:
                            bg_style = "background-color: #d4edda; color: #155724; font-weight: bold;"
                        elif curr_p < prev_p:
                            bg_style = "background-color: #f8d7da; color: #721c24; font-weight: bold;"
                        else:
                            bg_style = ""
                        styles.at[row_idx, col] = bg_style
                        if corresponding_vol_col in styles.columns:
                            styles.at[row_idx, corresponding_vol_col] = bg_style
                return styles

            styled_transposed_matrix = df_transposed_matrix.style.apply(style_transposed_cells, axis=None).format(precision=2)
            
            # Use the placeholder handle to safely update the container canvas space
            dashboard_history_placeholder.dataframe(styled_transposed_matrix, use_container_width=True, hide_index=True)
        else:
            dashboard_history_placeholder.info("Awaiting completed duration intervals to generate the historical matrix grid dashboard rows.")

        # Render Tab 6
        if not st.session_state.puvu_history.empty:
            logger.debug("Rendering Cumulative Price & Volume Gaining Breakouts tab.")
            puvu_placeholder.dataframe(st.session_state.puvu_history.style.apply(color_sr_signals, axis=1).format(precision=2), use_container_width=True, hide_index=True)
        else:
            puvu_placeholder.info("Awaiting the first completed pass to register Price Up & Volume Up signals.")

        # Render Tab 7
        if not st.session_state.pdvd_history.empty:
            logger.debug("Rendering Cumulative Price & Volume Dropping Breakdowns tab.")
            pdvd_placeholder.dataframe(st.session_state.pdvd_history.style.apply(color_sr_signals, axis=1).format(precision=2), use_container_width=True, hide_index=True)
        else:
            pdvd_placeholder.info("Awaiting the first completed pass to register Price Down & Volume Down signals.")

        # Render Tab 8 (PriceUp VolumeDown)
        if not st.session_state.puvd_history.empty:
            puvd_placeholder.dataframe(st.session_state.puvd_history.style.apply(color_sr_signals, axis=1).format(precision=2), use_container_width=True, hide_index=True)
        else:
            puvd_placeholder.info("Awaiting the first completed pass to register Price Up & Volume Down signals.")

        # Render Tab 9 (PriceDown VolumeUp)
        if not st.session_state.pdvu_history.empty:
            pdvu_placeholder.dataframe(st.session_state.pdvu_history.style.apply(color_sr_signals, axis=1).format(precision=2), use_container_width=True, hide_index=True)
        else:
            pdvu_placeholder.info("Awaiting the first completed pass to register Price Down & Volume Up signals.")

        # Frequency throttling barrier (Pausation check matching your few-seconds interval)
        time.sleep(3)