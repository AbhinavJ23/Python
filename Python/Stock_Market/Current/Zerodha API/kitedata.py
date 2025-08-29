import os,sys,time
from sys import platform
import pandas as pd
import xlwings as xw
import numpy as np
from datetime import datetime, timedelta
from baselogger import logger
import ctypes
from baselogin import Login
from kitemain import KiteMain

# Start - Call splash to show program loading
if getattr(sys, 'frozen', False):
    import pyi_splash

############################# Start - Function to check validity,expiry #############################
def check_validity():
    valid_from_str = '28/08/2025 00:00:00'
    valid_from_time = datetime.strptime(valid_from_str, '%d/%m/%Y %H:%M:%S')
    valid_till_time = valid_from_time + timedelta(days=30)
    time_now = datetime.now()
    time_left = valid_till_time - time_now
    logger.debug(f'Time Left - {time_left}')
    logger.debug(f'Days left - {time_left.days}')
    total_seconds = time_left.total_seconds()
    logger.debug(f'Total Seconds Left - {total_seconds}')
    if total_seconds < 0:
        if platform == "win32":
            ctypes.windll.user32.MessageBoxW(0, "Your product usage period has ended!", "Error",0)
        elif platform == "darwin":
            command_str = "osascript -e 'Tell application \"System Events\" to display dialog \"Your product usage period has ended!\" with title \"Error\"'"
            os.system(command_str)
        elif platform == "linux":
            logger.debug("Usage period error message not implemented for Linux yet")

        return False
    else:    
        hours = int((total_seconds - time_left.days*24*60*60)//3600)
        minutes = int((total_seconds - time_left.days*24*60*60 - hours*60*60)//60)
        seconds = round(total_seconds - time_left.days*24*60*60 - hours*60*60 - minutes*60)
        message = "Your product usage period will expire in " + str(time_left.days) + " day(s) " + str(hours) +" hours(s) " + str(minutes) + " min(s) and " + str(seconds) + " second(s)"
        if platform == "win32":
            ctypes.windll.user32.MessageBoxW(0, message, "Warning",0)
        elif platform == "darwin":
            command_str = "osascript -e 'Tell application \"System Events\" to display dialog \""+message+"\" with title \"Warning\"'"
            os.system(command_str)
        elif platform == "linux":
            logger.debug("Usage period warning message not implemented for Linux yet")

        return True
############################# End - Function to check validity,expiry #############################

# End - Call splash to show program loading
if getattr(sys, 'frozen', False):
    pyi_splash.close()

# Show login screen
login = Login()

# Check if login is successful
if login.is_logged_in:
    logger.debug("Logged in. Proceeding further....")
else:
    logger.error("Not Logged in. Exiting....")
    sys.exit()

# Check validity and expiry
status = check_validity()
if not status:
   sys.exit()

# All good. Proceeding with Kite Data
kite = KiteMain()

## Creating new excel and adding sheets
file_name = 'KiteData_'+time.strftime('%Y%m%d%H%M%S')+'.xlsx'
if not os.path.exists(file_name):
    try:
        wb = xw.Book()
        wb.sheets.add('PriceDownVolumeUp')
        wb.sheets.add('PriceUpVolumeDown')
        wb.sheets.add("PriceDownVolumeDown")
        wb.sheets.add("PriceUpVolumeUp")
        wb.sheets.add("MinCumulativeTurnover")
        wb.sheets.add("MaxCumulativeTurnover")
        #wb.sheets.add("SpotTurnover")
        #wb.sheets.add("SpotVolume")
        #wb.sheets.add("SpotPrice")
        #wb.sheets.add("FuturesData")
        wb.sheets.add("EquityData")
        #wb.sheets.add("OptionChain")
        wb.sheets.add("Configuration")
        wb.save(file_name)
        logger.debug("Created Excel - " + file_name)
    except Exception as e:
        logger.critical(f'Error Creating Excel - {e}')
        sys.exit()

wb = xw.Book(file_name)
cfg = wb.sheets("Configuration")
#oc = wb.sheets("OptionChain")
eq = wb.sheets("EquityData")
#fd = wb.sheets("FuturesData")
#sp = wb.sheets("SpotPrice")
#sv = wb.sheets("SpotVolume")
#st = wb.sheets("SpotTurnover")
maxct = wb.sheets("MaxCumulativeTurnover")
minct = wb.sheets("MinCumulativeTurnover")
puvu = wb.sheets("PriceUpVolumeUp")
pdvd = wb.sheets("PriceDownVolumeDown")
puvd = wb.sheets("PriceUpVolumeDown")
pdvu = wb.sheets("PriceDownVolumeUp")

####################### Initializing Constants #######################
COLOR_GREY = (211, 211, 211)
COLOR_GREEN = (0, 255, 0)
COLOR_DARK_GREEN = (0, 100, 0)
COLOR_RED = (255, 0, 0)
COLOR_YELLOW = (255, 255, 0)
COLOR_CYAN = (0, 255, 255)
MAX_VOLUME_PERCENT_DIFF = 500
MAX_TURNOVER_PERCENT_DIFF = 500
MAX_TURNOVER_VALUE_DIFF = 50000000
ONE_CRORE = 10000000
TEN_CRORE = 100000000
FIFTY_CRORE = 5000000000
HUNDRED_CRORE = 10000000000
DURATION = 5 #Mins
MAX_CUMULATIVE_TURNOVER = 100000000
MIN_CUMULATIVE_TURNOVER = 20000000
PERCENT_UP = 1 #Percent
PERCENT_DOWN = -1 #Percent
MARKET_OPEN_DURATION = 375 #Mins
DELIVERY_CHANGE_DURATION = 30 #Mins
RISK_FREE_INT_RATE = 5 #Percent
OPTION_PCR_DURATION = 5 #Mins

####################### Initializing Excel Sheets #######################
"""
oc.range('1:1').font.bold = True
oc.range('1:1').color = COLOR_GREY
oc.range('C1:C200').color = COLOR_GREY
oc.range('H1:H500').color = COLOR_GREY
oc.range('C1').column_width = 2
oc.range('H1').column_width = 2
fd.range('1:1').font.bold = True
fd.range('1:1').color = COLOR_GREY
fd.range('B1:C200').color = COLOR_GREY
fd.range('B1').column_width = 1
fd.range('C1').column_width = 1
fd.range('H1').column_width = 2
fd.range('H1:H1000').color = COLOR_GREY

"""
eq.range('1:1').font.bold = True
eq.range('1:1').color = COLOR_GREY
eq.range('B1:C40').color = COLOR_GREY
eq.range('B1').column_width = 1
eq.range('C1').column_width = 1
eq.range('H1').column_width = 2
eq.range('H1:H510').color = COLOR_GREY

######################### Initializing Configuration sheet #######################
cfg.range('D1').value = "IMPORTANT! KINDLY USE THIS SHEET TO CHECK CONFIGURATIONS AND OTHER DETAILS. ADD/MODIFY AS REQUIRED."
cfg.range('D1').font.bold = True
cfg.range('A1:CZ1').color = COLOR_GREY
cfg.range('A2').value = "EQUITY"
cfg.range('A2').font.bold = True
cfg.range('A2').color = COLOR_YELLOW
cfg.range('A3').value = "Max Cumulative Turnover"
cfg.range('A3').font.bold = True
cfg.range('B3').value = MAX_CUMULATIVE_TURNOVER/ONE_CRORE
cfg.range('C3').value = "₹ Cr"
cfg.range('A4').value = "Min Cumulative Turnover"
cfg.range('A4').font.bold = True
cfg.range("A3").autofit()
cfg.range('B4').value = MIN_CUMULATIVE_TURNOVER/ONE_CRORE
cfg.range('C4').value = "₹ Cr"
cfg.range('A5').value = "Duration"
cfg.range('A5').font.bold = True
cfg.range('B5').value = DURATION
cfg.range('C5').value = "Mins"
cfg.range('D2:D5').color = COLOR_GREY
cfg.range('E3').value = "Percent Up"
cfg.range('E3').font.bold = True
cfg.range('F3').value = PERCENT_UP
cfg.range('G3').value = "%"
cfg.range('E4').value = "Percent Down"
cfg.range('E4').font.bold = True
cfg.range("E4").autofit()
cfg.range('F4').value = PERCENT_DOWN
cfg.range('G4').value = "%"
cfg.range('H2:H5').color = COLOR_GREY
cfg.range('A6:CZ6').color = COLOR_GREY
cfg.range('A7').value = "Stock List"
cfg.range('A7').font.bold = True
cfg.range('A8').value = "Support"
cfg.range('A8').font.bold = True
cfg.range('A9').value = "Resistance"
cfg.range('A9').font.bold = True
cfg.range('A11:CZ11').color = COLOR_GREY
"""
cfg.range('A12').value = "OPTIONS"
cfg.range('A12').font.bold = True
cfg.range('A12').color = COLOR_YELLOW
cfg.range('A13').value = "Option PCR Duration"
cfg.range('A13').font.bold = True
#cfg.range("A13").autofit()
cfg.range('B13').value = OPTION_PCR_DURATION
cfg.range('C13').value = "Mins"
cfg.range('A14').value = "Risk Free Interest Rate"
cfg.range('A14').font.bold = True
cfg.range('B14').value = RISK_FREE_INT_RATE
cfg.range('C14').value = "%"
"""
logger.debug("Configurations sheet initialized")

####################### Initializing OptionChain sheet #######################
"""
oc.range("A:B").value = oc.range("D6:E19").value = oc.range("H1:V4000").value = None
oc_df = None
oc_df= pd.DataFrame({"FNO Symbol":["NIFTY", "BANKNIFTY"] + nse.equity_market_data("Securities in F&O", symbol_list=True)})
if oc_df is not None:
    oc_df = oc_df.set_index("FNO Symbol", drop=True)
    oc.range("A1").value = oc_df
else:
    logger.error(f'Error getting FNO symbols dataframe for Options Data')
    time.sleep(5)
    logger.debug("Trying to connect again...")
    nse = NSE()

oc.range("A1:A200").autofit()
oc.range("D2").value, oc.range("D3").value = "Enter Symbol ->", "Enter Expiry ->"
oc.range('D2').font.bold = True
oc.range('D3').font.bold = True
oc.range("D2:E3").autofit()
oc.range('A200:B200').color = COLOR_GREY
oc.range('D20:G20').color = COLOR_GREY
oc.range("D1").value = "Current Time"
oc.range("F21").value = "Maximum"
oc.range('F21').font.bold = True
oc.range('F21:G21').merge()
oc.range("E22").value = "PCR"
oc.range('E22').font.bold = True
oc.range("F22").value = "PE OI"
oc.range('F22').font.bold = True
oc.range("G22").value = "CE OI"
oc.range('G22').font.bold = True
pre_oc_sym = pre_oc_exp = None
exp_list = []
logger.debug("OptionChain sheet initialized")
"""

######################### Initializing EquityData sheet #######################
eq.range("A:A").value = eq.range("D5:H30").value = eq.range("I1:AE510").value = None
eq_df = pd.DataFrame({"Index Symbol":kite.equity_market_categories})
eq_df = eq_df.set_index("Index Symbol", drop=True)
eq.range("A1").value = eq_df
eq.range("A1:A50").autofit()
eq.range("D2").value, eq.range("D3").value = "Enter Index ->", "Enter Equity ->"
eq.range("E2").value = eq.range("A2").value
eq.range('D2').font.bold = True
eq.range('D3').font.bold = True
eq.range('A40:G40').color = COLOR_GREY
eq.range("D1").value = "Index Time"
eq.range("D1").autofit()
eq.range("F1").value = "Equity Time"
eq.range("F1").autofit()
eq.range("F2").value = "Index Value"
eq.range('F2').font.bold = True
#pre_ind_sym = pre_eq_sym = None
eq.range("F3").value = "Equity Value"
eq.range('F3').font.bold = True
eq.range("D2:E3").autofit()
logger.debug("EquityData sheet initialized")

"""
####################### Initializing FuturesData sheet #######################
fd.range("A:A").value = fd.range("G1:AD100").value = None
fd_df = None
fd_df= pd.DataFrame({"FNO Symbol":["NIFTY", "BANKNIFTY"] + nse.equity_market_data("Securities in F&O", symbol_list=True)})
if fd_df is not None:
    fd_df = fd_df.set_index("FNO Symbol", drop=True)
    fd.range("A1").value = fd_df
else:
    logger.error(f'Error getting FNO symbols dataframe for Futures Data')
    time.sleep(5)
    logger.debug("Trying to connect again...")
    nse = NSE()

fd.range("A1:A200").autofit()
fd.range('A200:B200').color = COLOR_GREY
fd.range("D1").value = "Current Time"
fd.range("D1").autofit()
fd.range("F1").value = "Underlying Value"
fd.range("F1").autofit()
fd.range("D2").value = "Enter Symbol ->"
fd.range('D2').font.bold = True
fd.range("D2").autofit()
pre_fd_sym = None
logger.debug("FuturesData sheet initialized")

"""
logger.debug("All Excel sheets initialized")

####################### Initializing Global Variables #######################
row_number = 1
row_number_1 = 2
col_number = 2
col_number_1 = 1
prev_time = curr_time = None
prev_time_1 = datetime.now()
prev_time_2 = None
prev_vol = curr_vol = []
prev_vol_diff = curr_vol_diff = []
prev_price = curr_price = []
prev_price_diff = curr_price_diff = []
prev_turn = curr_turn = []
prev_turn_diff = curr_turn_diff = []
cum_turn_dict = {}
cum_price_diff_dict = {}
prev_cum_price_diff_dict = {}
cum_vol_diff_dict = {}
prev_cum_vol_diff_dict = {}
prev_min_cum_turn_dict = {}
price_vol_dict_flag = False
prev_eq_df = None
initial_len_eq_df = 0

"""
option_prev_time = None
option_curr_time = None
option_row_number = 1
option_duration = None
prev_pcr = None
pcr = None
#pcr_print_flag = False

"""
#price_vol_dir_supp_res_df = None #pd.DataFrame(columns=['PUVU_RES','PUVU_SUP','PDVD_RES','PDVD_SUP'])

############################# Start - Function to get excel column(A1,B1 etc) given a positive number #############################
alpha = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
def get_col_name(num):
    if num < 26:
        return alpha[num-1]
    else:
        q, r = num//26, num % 26
        if r == 0:
            if q == 1:
                return alpha[r-1]
            else:
                return get_col_name(q-1) + alpha[r-1]
        else:
            return get_col_name(q) + alpha[r-1]
############################# End - Function to get excel column(A1,B1 etc) given a positive integer #############################
"""
############################# Start - Function to get option greeks #############################
def get_option_greeks(df, call_or_put, expiry):
    logger.debug('Function get_option_greeks')
    time = ((datetime(expiry.year, expiry.month, expiry.day, 15, 30) - datetime.now()) / timedelta(days=1)) / 365
    logger.debug (f'Time to Expiry in Years - {time}')
    int_rate = RISK_FREE_INT_RATE
    greek_list = []
    try:
        for i, row in df.iterrows():
            strike = i            
            if call_or_put == 'c':
                last_price = row['CE LTP']
                imp_vol = row['CE IV']
            elif call_or_put == 'p':
                last_price = row['PE LTP']
                imp_vol = row['PE IV']
            if last_price <= 0 or imp_vol <= 0:
                 greek_list.append({"Delta": 0, "Gamma": 0, "Theta": 0, "Vega": 0, "Rho": 0})
            else:
                Delta = delta(call_or_put, last_price, strike, time, int_rate, imp_vol)
                Gamma = gamma(call_or_put, last_price, strike, time, int_rate, imp_vol)
                Theta = theta(call_or_put, last_price, strike, time, int_rate, imp_vol)
                Vega = vega(call_or_put, last_price, strike, time, int_rate, imp_vol)
                Rho = rho(call_or_put, last_price, strike, time, int_rate, imp_vol)
                greek_list.append({"Delta": Delta, "Gamma": Gamma, "Theta": Theta, "Vega": Vega, "Rho": Rho})
    except Exception as e:
        logger.error(f'Error getting Option Greeks - {e}')
        empty_df = pd.DataFrame(index=df.index, columns=['Delta','Gamma','Theta','Vega','Rho'])
        empty_df.fillna(0)
        return empty_df
    
    greek_df = pd.DataFrame(greek_list, index=df.index)
    return greek_df
############################# End - Function to get option greeks #############################

############################# Start - Function to get option OI #############################
def get_option_oi(df, call_or_put, spot_ltp):
    logger.debug('Function get_option_oi')
    oi = None
    for i, row in df.iterrows():
        strike = i
        if strike > spot_ltp:
            if call_or_put == 'c':
                oi = row['CE OI']
            elif call_or_put == 'p':
                oi = row['PE OI']
            break
        else:
            logger.debug(f'Function get_option_oi - ignoring strike {strike}')
    return oi
############################# End - Function to get option OI #############################

############################# Start - Function to get delivery info ###########################
def get_delivery_info(df):
    logger.debug('Function get_delivery_info')
    delivery_info_list = []
    flag = True
    for i in df.index:
        symbol = i
        data = None
        #if symbol != 'NIFTY 50':
        if 'NIFTY' not in symbol:
            data = nse.equity_info(symbol, trade_info=True)
            if data is not None:
                for key,value in data.items():              
                    if str(key) == "securityWiseDP":
                        delivery_info_list.append(value)
            else:
                logger.error(f'Error getting Delivery Info for {symbol} - Delivery Info Data is Null')
                flag = False
                break
        else:
            delivery_info_list.append({"quantityTraded": 'NA', "deliveryQuantity": 'NA', "deliveryToTradedQuantity": 'NA', "secWiseDelPosDate": 'NA'})
            logger.debug(f'Adding NA to {symbol} data')
    
    if len(delivery_info_list) == len(df.index) and flag:
        delivery_info_df = pd.DataFrame(delivery_info_list, index=df.index)
        delivery_info_df.drop(["seriesRemarks"],axis=1,inplace=True)
        logger.debug('Returning delivery_info_df')
        return delivery_info_df
    else:
        empty_df = pd.DataFrame(index=df.index, columns=['quantityTraded','deliveryQuantity','deliveryToTradedQuantity','secWiseDelPosDate'])
        empty_df.fillna(0)
        logger.debug('Returning empty_df')
        return empty_df

############################# End - Function to get delivery info #############################
"""
############################# Start - Function to get Equity Config DataFrame #############################
def get_equity_config_df():
    stocks_res_sup_df = None
    logger.debug("Getting Equity Config DataFrame")
    config_stock_list = cfg.range('B7:CZ7').value
    config_support_list = cfg.range('B8:CZ8').value
    config_resistance_list = cfg.range('B9:CZ9').value
    if config_stock_list:
        while None in config_stock_list:
            config_stock_list.remove(None)
    if config_stock_list and config_support_list and config_resistance_list:
        len_config_stock_list = len(config_stock_list)        
        config_support_list = config_support_list[0:len_config_stock_list]        
        config_resistance_list = config_resistance_list[0:len_config_stock_list]        
        logger.debug(f'Config stock list - {config_stock_list}')
        logger.debug(f'Config support list - {config_support_list}')
        logger.debug(f'Config resistance list - {config_resistance_list}')
        stocks_res_sup_df = pd.DataFrame(list(zip(config_support_list, config_resistance_list)), index=config_stock_list, columns = ['Support','Resistance'])
        logger.debug(f'Printing stocks df - {stocks_res_sup_df}')
    return stocks_res_sup_df
############################# End - Function to get Equity Config DataFrame #############################

############################# Start - Function to create Spot sheets #############################
def create_spot_data(df,type,time,duration,row_number,prev_spot,curr_spot,prev_spot_diff,curr_spot_diff,col_number_1,stock_list):
    logger.debug(f'Creating Spot data - {type} for {time} and row {row_number}')
    if type == "Price":
        spot_df = df[["last_price"]]
        #sh = sp
    elif type == "Volume":
        spot_df = df[["volume"]]
        #sh = sv
    elif type == "Turnover":
        spot_df = df[["turnover"]]
        spot_df_price = df[["last_price"]]
        #sh = st
    else:
        logger.error(f'Error! Unexpected Input - {type}')
        return

    #if row_number == 1:
    #    sh.range("A1:ZZ1").color = COLOR_GREY

    spot_df1 = spot_df.transpose()
    col_number = 2
    iter = 0
    per_diff = 0    
    vol_list = []
    turn_list = []
    global cum_price_diff_dict
    global cum_vol_diff_dict
    global price_vol_dict_flag
    for col_name in spot_df1:
        if col_name not in kite.equity_market_categories:
            spot_value = spot_df1[col_name].values
            if spot_value is None:
                logger.debug(f'For Time {time} - In {type} data for {col_name}, value is None, hence setting it to zero')
                spot_value = np.zeros(1)
            if row_number == 1:
                """sh.range(f'{get_col_name(col_number)}' + str(row_number)).options(index=False).value = spot_df1[col_name]            
                sh.range(f'{get_col_name(col_number+1)}' + str(row_number)).value = "Difference"
                if type in ("Volume","Turnover"):
                    sh.range(f'{get_col_name(col_number+2)}' + str(row_number)).value = "% Change"
                if iter == 0:
                    sh.range(f'A{row_number + 1}').value = time
                    sh.range(f'A{row_number + 1}').font.bold = True"""
                prev_spot.append(spot_value)
            else:
                """if iter == 0:                  
                    sh.range(f'A{row_number + 1}').value = time.strftime('%H:%M:%S')
                    sh.range(f'A{row_number + 1}').font.bold = True
                sh.range(f'{get_col_name(col_number)}'+ str(row_number+1)).value = spot_value"""
                curr_spot.append(spot_value)
                if curr_spot[iter] is not None and prev_spot[iter] is not None:
                    val_diff = curr_spot[iter] - prev_spot[iter]
                else:
                    logger.debug(f'For Time {time} - In {type} data for {col_name},  current value is {curr_spot[iter]} and previous value is {prev_spot[iter]}, hence setting its difference to zero')
                    val_diff = np.zeros(1)
                if type in ("Volume","Turnover") and val_diff < 0:
                    ## This could be due to some data issue
                    ## It could also be possible that value (price/volume/turover) has not changed since last time.
                    logger.warning(f'Potential Data Issue in {type} data!! value diff is {val_diff}')
                    logger.debug(f'For Time {time} - In {type} data for {col_name}, current value {curr_spot[iter]} can not be less than previous value {prev_spot[iter]}. Hence setting difference between current and previous value as zero')
                    val_diff = np.zeros(1)
                #sh.range(f'{get_col_name(col_number+1)}'+ str(row_number+1)).value = val_diff

                ## First time add to Cumulative Price, Volume difference
                if not price_vol_dict_flag:
                    if type == "Price":
                        cum_price_diff_dict.update({col_name:val_diff})
                    elif type == "Volume":
                        cum_vol_diff_dict.update({col_name:val_diff})
                if row_number == 2:
                    prev_spot_diff.append(val_diff)
                    """if type == "Price":
                        if val_diff > 0:
                            sh.range(f'{get_col_name(col_number+1)}'+ str(row_number+1)).color = COLOR_GREEN
                        elif val_diff < 0:
                            sh.range(f'{get_col_name(col_number+1)}'+ str(row_number+1)).color = COLOR_RED"""
                else:
                    curr_spot_diff.append(val_diff)
                    if type in ("Volume","Turnover"):
                        if prev_spot_diff[iter] != 0:
                            ## Calculate percentage difference only if denominator is not zero
                            per_diff = ((curr_spot_diff[iter] - prev_spot_diff[iter])*100)/prev_spot_diff[iter]
                            #sh.range(f'{get_col_name(col_number+2)}'+ str(row_number+1)).value = per_diff
                        else:
                            logger.warning("Division by Zero occurred!!")
                            logger.debug(f'For Time {time} - Avoiding division by zero in {type} data for {col_name}')
                            #sh.range(f'{get_col_name(col_number+2)}'+ str(row_number+1)).value = "NA"
                            per_diff = np.zeros(1)
                            #sh.range(f'{get_col_name(col_number+2)}'+ str(row_number+1)).color = COLOR_YELLOW

                    if type == "Volume":
                        #if per_diff > MAX_VOLUME_PERCENT_DIFF:
                        #    sh.range(f'{get_col_name(col_number+2)}'+ str(row_number+1)).color = COLOR_GREEN
                        #elif per_diff < -MAX_VOLUME_PERCENT_DIFF:
                        #    sh.range(f'{get_col_name(col_number+2)}'+ str(row_number+1)).color = COLOR_RED

                        #if per_diff > MAX_VOLUME_PERCENT_DIFF or per_diff < -MAX_VOLUME_PERCENT_DIFF:
                        #if col_name not in kite.equity_market_categories:
                        stock_list.append(col_name)
                        vol_list.append(per_diff)
                        #else:
                            #logger.debug("Ignorning Stock Name - " + col_name)
                        ## Logic for updating cumulative volume difference
                        if price_vol_dict_flag:
                            cum_vol_diff_dict[col_name] = cum_vol_diff_dict[col_name] + val_diff
                    elif type == "Turnover":
                        #if per_diff > MAX_TURNOVER_PERCENT_DIFF:
                        #    sh.range(f'{get_col_name(col_number+2)}'+ str(row_number+1)).color = COLOR_GREEN
                        #elif per_diff < -MAX_TURNOVER_PERCENT_DIFF:
                        #    sh.range(f'{get_col_name(col_number+2)}'+ str(row_number+1)).color = COLOR_RED
                        temp_val_diff = val_diff/ONE_CRORE
                        if col_name in stock_list:
                            turn_list.append(temp_val_diff)
                        else:
                            logger.warning("Not Expected!! Stock Name - " + col_name)
                        #elif col_name not in kite.equity_market_categories:
                        #if val_diff > MAX_TURNOVER_VALUE_DIFF or per_diff > MAX_TURNOVER_PERCENT_DIFF or per_diff < -MAX_TURNOVER_PERCENT_DIFF:
                            #stock_list.append(col_name)
                            #turn_list.append(temp_val_diff)
                    elif type == "Price":
                        #if val_diff > 0:
                        #    sh.range(f'{get_col_name(col_number+1)}'+ str(row_number+1)).color = COLOR_GREEN
                        #elif val_diff < 0:
                        #    sh.range(f'{get_col_name(col_number+1)}'+ str(row_number+1)).color = COLOR_RED
                        ## Logic for updating cumulative price difference
                        if price_vol_dict_flag:
                            cum_price_diff_dict[col_name] = cum_price_diff_dict[col_name] + val_diff             
            iter += 1
            col_number += 3
        else:
            logger.debug(f'Ignoring {col_name} in {type} data as it is an Index Symbol')
    #if row_number == 1:
    #    sh.range("A1:ZZ1").font.bold = True
    #    sh.range("A1:ZZ1").autofit()
    #if type in ("Volume","Turnover"):
    if type in ("Turnover"):
        create_cum_turn_sheet(type,time,duration,row_number,col_number_1,turn_list,stock_list, spot_df_price)
        #create_cum_turn_sheet(type,time,duration,row_number,col_number_1,vol_list,turn_list,stock_list)
############################# End - Function to create Spot sheets #############################

############################# Start - Function to creete Cumulative turnover data ############
def create_cum_turn_sheet(type,time,duration,row_number,col_number_1,turn_list,stock_list, spot_df_price):
    logger.debug(f'Printing CumulativeTurnover Sheet for {time} and row {row_number}')
    global cum_turn_dict
    if row_number >= 2:
        #maxct.range(f'{get_col_name(col_number_1)}' + '1').value = time.strftime("%H:%M:%S")
        #maxct.range(f'{get_col_name(col_number_1)}' + '1').font.bold = True

        """if type == "Volume":
            if stock_list:
                logger.debug(f'For MaxVolume, Stock List is - {stock_list}')
                logger.debug(f'For MaxVolume, Volume List is - {vol_list}')
                temp_vol_df = pd.DataFrame(vol_list, index=stock_list, columns=['Vol%Diff'])
                maxct.range(f'{get_col_name(col_number_1)}' + '2').value = temp_vol_df
            else:
                maxct.range(f'{get_col_name(col_number_1+1)}' + '2').value = "Vol%Diff"
            maxct.range(f'{get_col_name(col_number_1)}' + '2').value = "Name"
            maxct.range(f'{get_col_name(col_number_1)}' + '2').font.bold = True
            maxct.range(f'{get_col_name(col_number_1+1)}' + '2').font.bold = True            
            logger.debug("MaxVolume Printed")"""
        #elif type == "Turnover":
        if stock_list and turn_list:
            #stock_list.sort()
            logger.debug(f'For MaxTurnover, Stock List is - {stock_list}')
            logger.debug(f'For MaxTurnover, Turnover List is - {turn_list}')
            #temp_turn_df = pd.DataFrame(turn_list, index=stock_list, columns=['Turnover(₹ Cr)'])
            #maxct.range(f'{get_col_name(col_number_1+2)}' + '2').value = temp_turn_df
            #if (temp_turn_df['Turnover(₹ Cr)'] > FIFTY_CRORE/ONE_CRORE).any():
                #logger.debug("Max Turnover greater than 50 Cr found")
                #maxct.range(f'{get_col_name(col_number_1+2)}' + '2').color = COLOR_GREEN
                #maxct.range(f'{get_col_name(col_number_1+3)}' + '2').color = COLOR_GREEN

            if not cum_turn_dict:
                cum_turn_dict = {stock_list[i]: turn_list[i] for i in range(len(stock_list))}
            else:
                temp_cum_turn_dict = {stock_list[i]: turn_list[i] for i in range(len(stock_list))}
                for key in temp_cum_turn_dict:
                    if key in cum_turn_dict:
                        cum_turn_dict[key] = cum_turn_dict[key] + temp_cum_turn_dict[key]
                    else:
                        cum_turn_dict.update({key:temp_cum_turn_dict[key]})
            logger.debug(f'Intermediate cumulative turnover - {cum_turn_dict}')
        #else:
            #maxct.range(f'{get_col_name(col_number_1+3)}' + '2').value = "Turnover(₹ Cr)"
        #maxct.range(f'{get_col_name(col_number_1+2)}' + '2').value = "Name"
        #maxct.range(f'{get_col_name(col_number_1+2)}' + '2').font.bold = True
        #maxct.range(f'{get_col_name(col_number_1+3)}' + '2').font.bold = True
        #maxct.range(f'{get_col_name(col_number_1+3)}' + '2').autofit()
        #logger.debug("MaxTurnover Printed")
            
        if duration.total_seconds()/60 >= DURATION:
            create_cum_turn_sheet_min_max("Max",time, col_number_1, cum_turn_dict, spot_df_price)
            create_cum_turn_sheet_min_max("Min",time, col_number_1, cum_turn_dict, spot_df_price)
            #maxct.range(f'{get_col_name(col_number_1)}' + '1').value = time.strftime("%H:%M:%S")
            #maxct.range(f'{get_col_name(col_number_1)}' + '1').font.bold = True
            #maxct.range(f'{get_col_name(col_number_1)}' + '1').autofit()
            #maxct.range(f'{get_col_name(col_number_1+4)}' + '1').value = "Cumulative Turnover"
            #maxct.range(f'{get_col_name(col_number_1+4)}' + '1').font.bold = True
            #maxct.range(f'{get_col_name(col_number_1+4)}'+ '1').color = COLOR_YELLOW
            #maxct.range(f'{get_col_name(col_number_1+5)}'+ '1').color = COLOR_YELLOW
            #logger.debug(f'Cumulative turnover Before - {cum_turn_dict}')
            #cum_turn_dict_max = {key: cum_turn_dict[key] for key in cum_turn_dict if cum_turn_dict[key] >= MAX_CUMULATIVE_TURNOVER/ONE_CRORE}
            #logger.debug(f'Max Cumulative turnover After - {cum_turn_dict_max}')
            #sorted_cum_turn_dict_max = dict(sorted(cum_turn_dict_max.items()))
            #cum_turn_df_max = pd.DataFrame(sorted_cum_turn_dict_max.values(), index=sorted_cum_turn_dict_max.keys(), columns=['Turnover(₹ Cr)'])
            #maxct.range(f'{get_col_name(col_number_1)}' + '2').value = cum_turn_df_max
            #maxct.range(f'{get_col_name(col_number_1)}' + '2').value = "Name"
            #maxct.range(f'{get_col_name(col_number_1)}' + '2').font.bold = True
            #maxct.range(f'{get_col_name(col_number_1 + 1)}' + '2').font.bold = True
            #maxct.range(f'{get_col_name(col_number_1)}' + '2').autofit()
            #maxct.range(f'{get_col_name(col_number_1 + 1)}' + '2').autofit()
            #logger.debug('Cumulative turonver printed')
            #check_price_up_down(col_number_1, sorted_cum_turn_dict, prev_cum_price_diff_dict, cum_price_diff_dict)
############################# End - Function to creete Cumulative turnover data ############

############################# Start - Function to print Min and Max Cumulative turnover ############
def create_cum_turn_sheet_min_max(type, time, col_number_1, cum_turn_dict, spot_df_price):
    logger.debug(f'Printing {type}CumulativeTurnover Sheet for {time}')
    logger.debug(f'{type}CumulativeTurnover Before - {cum_turn_dict}')
    sh = None
    global prev_min_cum_turn_dict
    if type == "Max":
        sh = maxct
        temp_cum_turn_dict = {key: cum_turn_dict[key] for key in cum_turn_dict if cum_turn_dict[key] >= MAX_CUMULATIVE_TURNOVER/ONE_CRORE}
    elif type == "Min":
        sh = minct
        temp_cum_turn_dict = {key: cum_turn_dict[key] for key in cum_turn_dict if cum_turn_dict[key] <= MIN_CUMULATIVE_TURNOVER/ONE_CRORE}
        prev_min_cum_turn_dict = temp_cum_turn_dict
    else:
        logger.error(f'Error! Unexpected Input - {type}')
        return
    logger.debug(f'{type}CumulativeTurnover After - {temp_cum_turn_dict}')
    sh.range(f'{get_col_name(col_number_1)}' + '1').value = time.strftime("%H:%M:%S")
    sh.range(f'{get_col_name(col_number_1)}' + '1').font.bold = True
    sh.range(f'{get_col_name(col_number_1)}' + '1').autofit()
    sorted_cum_turn_dict = dict(sorted(temp_cum_turn_dict.items()))
    temp_cum_turn_df = pd.DataFrame(sorted_cum_turn_dict.values(), index=sorted_cum_turn_dict.keys(), columns=['Turnover(₹ Cr)'])
    sh.range(f'{get_col_name(col_number_1)}' + '2').value = temp_cum_turn_df
    sh.range(f'{get_col_name(col_number_1)}' + '2').value = "Name"
    sh.range(f'{get_col_name(col_number_1)}' + '2').font.bold = True
    sh.range(f'{get_col_name(col_number_1 + 1)}' + '2').font.bold = True
    #sh.range(f'{get_col_name(col_number_1)}' + '2').autofit()
    sh.range(f'{get_col_name(col_number_1 + 1)}' + '2').autofit()
    # Logic to print price next to cumulative turnover
    temp_price_list = []
    if temp_cum_turn_df is not None:
        for name in temp_cum_turn_df.index:
            temp_price_list.append(spot_df_price.loc[name, 'last_price'])
        temp_price_df = pd.DataFrame(temp_price_list, index=temp_cum_turn_df.index, columns=['Price'])
    else:
        logger.error(f'Error! temp_cum_turn_df is None for {type}CumulativeTurnover')
        return
    sh.range(f'{get_col_name(col_number_1 + 2)}' + '2').options(index=False).value = temp_price_df
    sh.range(f'{get_col_name(col_number_1 + 2)}' + '2').font.bold = True
    logger.debug(f'{type}CumulativeTuronver printed')
    check_price_up_down(type, col_number_1, sorted_cum_turn_dict, prev_cum_price_diff_dict, cum_price_diff_dict, prev_min_cum_turn_dict)
############################# End - Function to print Min and Max Cumulative turnover ############

############################# Start - Function to check if cumulative turnover stock price ############
############################# is up or down compared to previous price #############################
def check_price_up_down(type, col_number_1, turn_dict, prev_price_dict, curr_price_dict, prev_min_cum_turn_dict):
    logger.debug(f'Checking Price Up Down for {type}CumulativeTurnover')
    logger.debug(f'Previous Cumulative Price Difference Dict - {prev_price_dict}')
    logger.debug(f'Current Cumulative Price Difference Dict - {curr_price_dict}')
    sh = None
    global prev_eq_df
    if type == "Max":
        sh = maxct
    elif type == "Min":
        sh = minct
    else:
        logger.error(f'Error! Unexpected Input - {type}')
        return
    if prev_price_dict and curr_price_dict and prev_eq_df is not None:
        counter = 2
        for key in turn_dict:
            if key in prev_price_dict and key in curr_price_dict and key in prev_eq_df.index:
                if curr_price_dict[key] > prev_eq_df.loc[key, 'last_price'] * PERCENT_UP/100:
                    logger.debug(f'Current Equity price of {key} is {PERCENT_UP}% greater than last price')
                    sh.range(f'{get_col_name(col_number_1)}'+ str(counter + 1)).color = COLOR_GREEN
                    ## Checking if any stock from prev min cumulative turnover has moved to max cumulative turnover
                    if type == "Max" and key in prev_min_cum_turn_dict:
                        logger.debug(f'{key} has moved to MaxCumulativeTurnover from MinCumulativeTurnover with Price increase')
                        sh.range(f'{get_col_name(col_number_1)}'+ str(counter + 1)).color = COLOR_DARK_GREEN
                elif curr_price_dict[key] < prev_eq_df.loc[key, 'last_price'] * PERCENT_DOWN/100:
                    logger.debug(f'Current Equity price of {key} is {PERCENT_DOWN}% less than last price')
                    sh.range(f'{get_col_name(col_number_1)}'+ str(counter + 1)).color = COLOR_RED
                elif curr_price_dict[key] == 0:
                    logger.debug(f'Current Equity price of {key} is equal to last price')
                    sh.range(f'{get_col_name(col_number_1)}'+ str(counter + 1)).color = COLOR_YELLOW
                else:
                    logger.debug(f'Current Equity price of {key} is mostly neutral')
            else:
                logger.warning(f'Not Expected!! Key {key} not found in previous or current cumulative price difference dicts or prev_eq_df')
            counter += 1
############################# End - Function to check if cumulative turnover stock price ############
############################# is up or down compared to previous price #############################

def validate_buy_sell(df1, df2, list1, list2):
    for idx in df2.index:
        if idx in list1:
            if df1.loc[idx, 'last_price'] >= df2.loc[idx, 'PUVU_RES']:
                logger.debug(f'Price Up, Volume Up - Current Equity price of {idx} is greater than or equal to last price')
            elif df1.loc[idx, 'last_price'] <= df2.loc[idx, 'PUVU_SUP']:
                logger.debug(f'Price Up, Volume Up - Current Equity price of {idx} is less than or equal to last price')
        elif idx in list2:
            if df1.loc[idx, 'last_price'] >= df2.loc[idx, 'PDVD_RES']:
                logger.debug(f'Price Down, Volume Down - Current Equity price of {idx} is greater than or equal to last price')
            elif df1.loc[idx, 'last_price'] <= df2.loc[idx, 'PDVD_SUP']:
                logger.debug(f'Price Down, Volume Down - Current Equity price of {idx} is less than or equal to last price')
        else:
            logger.debug('validate_buy_sell : Not Expected to reach here')
        
############################# Start - Function to compare Cumulative Price, Volume difference after every fixed duration ############
############################# Also print the direction (Up,Down) for each stock #########################################
def create_price_vol_sheets(df, time, row_number_1):
    logger.debug(f'Printing PriceVolumeDirection Sheets for {time} and row {row_number_1}')
    logger.debug(f'prev_cum_price_diff_dict - {prev_cum_price_diff_dict} and prev_cum_vol_diff_dict - {prev_cum_vol_diff_dict}')
    logger.debug(f'cum_price_diff_dict - {cum_price_diff_dict} and cum_vol_diff_dict - {cum_vol_diff_dict}')
    col_number_2 = 2
    price_vol_up_list = []
    price_down_vol_up_list = []
    price_vol_down_list = []
    price_up_vol_down_list = []
    #global price_vol_dir_supp_res_df
    row_difference = int(MARKET_OPEN_DURATION/DURATION) + 2
    temp_row_number = row_number_1 + row_difference
    if row_number_1 == 2:
        puvu.range(f'A{temp_row_number}' + ':' + f'DZ{temp_row_number}').color = COLOR_GREY
        puvu.range(f'A{temp_row_number + row_difference}' + ':' + f'DZ{temp_row_number + row_difference}').color = COLOR_GREY
        pdvd.range(f'A{temp_row_number}' + ':' + f'DZ{temp_row_number}').color = COLOR_GREY
        puvd.range(f'A{temp_row_number}' + ':' + f'DZ{temp_row_number}').color = COLOR_GREY
        pdvu.range(f'A{temp_row_number}' + ':' + f'DZ{temp_row_number}').color = COLOR_GREY
        puvu.range(f'A{row_number_1}').value = time
        #puvu.range(f'A{row_number_1}').font.bold = True
        puvu.range(f'A{temp_row_number+1}').value = "Price Up,Volume Up List"
        puvu.range(f'A{temp_row_number + row_difference+1}').value = "Filtered Stocks - As per Price Up Volume Up and Configuration"
        pdvd.range(f'A{row_number_1 - 1}').value = "Price Down,Volume Down List"
        pdvd.range(f'A{temp_row_number+1}').value = "Filtered Stocks - As per Price Down Volume Down and Configuration"
        puvd.range(f'A{row_number_1 - 1}').value = "Price Up,Volume Down List"
        puvd.range(f'A{temp_row_number+1}').value = "Filtered Stocks - As per Price Up Volume Down and Configuration"
        pdvu.range(f'A{row_number_1 - 1}').value = "Price Down,Volume Up List"
        pdvu.range(f'A{temp_row_number+1}').value = "Filtered Stocks - As per Price Down Volume Up and Configuration"
    else:
        puvu.range(f'A{row_number_1}').value = time.strftime('%H:%M:%S')
        #puvu.range(f'A{row_number_1}').font.bold = True
        puvu.range(f'A{temp_row_number + 1}').value = time.strftime('%H:%M:%S')
        puvu.range(f'A{temp_row_number + row_difference + 1}').value = time.strftime('%H:%M:%S')
        pdvd.range(f'A{row_number_1 - 1}').value = time.strftime('%H:%M:%S')
        pdvd.range(f'A{temp_row_number + 1}').value = time.strftime('%H:%M:%S')
        puvd.range(f'A{row_number_1 - 1}').value = time.strftime('%H:%M:%S')
        puvd.range(f'A{temp_row_number + 1}').value = time.strftime('%H:%M:%S')
        pdvu.range(f'A{row_number_1 - 1}').value = time.strftime('%H:%M:%S')
        pdvu.range(f'A{temp_row_number + 1}').value = time.strftime('%H:%M:%S')

    puvu.range(f'A{row_number_1}').font.bold = True
    puvu.range(f'A{temp_row_number + 1}').font.bold = True
    puvu.range(f'A{temp_row_number + row_difference + 1}').font.bold = True
    pdvd.range(f'A{row_number_1 - 1}').font.bold = True
    pdvd.range(f'A{temp_row_number + 1}').font.bold = True
    puvd.range(f'A{row_number_1 - 1}').font.bold = True
    puvd.range(f'A{temp_row_number + 1}').font.bold = True
    pdvu.range(f'A{row_number_1 - 1}').font.bold = True
    pdvu.range(f'A{temp_row_number + 1}').font.bold = True

    for key in cum_price_diff_dict:
        if row_number_1 == 2:       
            ## Initializing the stock names and columns
            puvu.range(f'{get_col_name(col_number_2)}' + str(row_number_1-1)).value = key
            puvu.range(f'{get_col_name(col_number_2)}' + str(row_number_1-1)).autofit()
            puvu.range(f'{get_col_name(col_number_2)}' + str(row_number_1-1)).font.bold = True
            puvu.range(f'{get_col_name(col_number_2)}' + str(row_number_1)).value = "Price"
            puvu.range(f'{get_col_name(col_number_2)}' + str(row_number_1)).font.bold = True
            puvu.range(f'{get_col_name(col_number_2+1)}' + str(row_number_1)).value = "Volume"
            puvu.range(f'{get_col_name(col_number_2+1)}' + str(row_number_1)).font.bold = True
        else:
            ## Check difference between previous cumulative volume and current cumulative volume along with current cumulative price
            #if (cum_price_diff_dict[key] - prev_cum_price_diff_dict[key]) == 0:
            if cum_price_diff_dict[key] == 0:
                puvu.range(f'{get_col_name(col_number_2)}' + str(row_number_1)).value = "Neutral"
                puvu.range(f'{get_col_name(col_number_2)}' + str(row_number_1)).color = COLOR_YELLOW
                if (cum_vol_diff_dict[key] - prev_cum_vol_diff_dict[key]) == 0:
                    puvu.range(f'{get_col_name(col_number_2+1)}' + str(row_number_1)).value = "Neutral"
                    puvu.range(f'{get_col_name(col_number_2+1)}' + str(row_number_1)).color = COLOR_YELLOW
                elif (cum_vol_diff_dict[key] - prev_cum_vol_diff_dict[key]) > 0:
                    puvu.range(f'{get_col_name(col_number_2+1)}' + str(row_number_1)).value = "Up"
                    puvu.range(f'{get_col_name(col_number_2+1)}' + str(row_number_1)).color = COLOR_GREEN
                else:
                    puvu.range(f'{get_col_name(col_number_2+1)}' + str(row_number_1)).value = "Down"
                    puvu.range(f'{get_col_name(col_number_2+1)}' + str(row_number_1)).color = COLOR_RED      
            elif (cum_vol_diff_dict[key] - prev_cum_vol_diff_dict[key]) == 0:
                if cum_price_diff_dict[key] > 0:
                    puvu.range(f'{get_col_name(col_number_2)}' + str(row_number_1)).value = "Up"
                    puvu.range(f'{get_col_name(col_number_2)}' + str(row_number_1)).color = COLOR_GREEN
                else:
                    puvu.range(f'{get_col_name(col_number_2)}' + str(row_number_1)).value = "Down"
                    puvu.range(f'{get_col_name(col_number_2)}' + str(row_number_1)).color = COLOR_RED                  
                puvu.range(f'{get_col_name(col_number_2+1)}' + str(row_number_1)).value = "Neutral"
                puvu.range(f'{get_col_name(col_number_2+1)}' + str(row_number_1)).color = COLOR_YELLOW
            elif cum_price_diff_dict[key] < 0 and (cum_vol_diff_dict[key] - prev_cum_vol_diff_dict[key]) < 0:
                puvu.range(f'{get_col_name(col_number_2)}' + str(row_number_1)).value = "Down"
                puvu.range(f'{get_col_name(col_number_2)}' + str(row_number_1)).color = COLOR_RED
                puvu.range(f'{get_col_name(col_number_2+1)}' + str(row_number_1)).value = "Down"
                puvu.range(f'{get_col_name(col_number_2+1)}' + str(row_number_1)).color = COLOR_RED
                #Building Data for Price Down, Volume Down
                price_vol_down_list.append(key)
            elif cum_price_diff_dict[key] < 0 and (cum_vol_diff_dict[key] - prev_cum_vol_diff_dict[key]) > 0:
                puvu.range(f'{get_col_name(col_number_2)}' + str(row_number_1)).value = "Down"
                puvu.range(f'{get_col_name(col_number_2)}' + str(row_number_1)).color = COLOR_RED
                puvu.range(f'{get_col_name(col_number_2+1)}' + str(row_number_1)).value = "Up"
                puvu.range(f'{get_col_name(col_number_2+1)}' + str(row_number_1)).color = COLOR_GREEN
                #Building Data for Price Down, Volume Up
                price_down_vol_up_list.append(key)
            elif cum_price_diff_dict[key] > 0 and (cum_vol_diff_dict[key] - prev_cum_vol_diff_dict[key]) > 0:
                puvu.range(f'{get_col_name(col_number_2)}' + str(row_number_1)).value = "Up"
                puvu.range(f'{get_col_name(col_number_2)}' + str(row_number_1)).color = COLOR_CYAN
                puvu.range(f'{get_col_name(col_number_2+1)}' + str(row_number_1)).value = "Up"
                puvu.range(f'{get_col_name(col_number_2+1)}' + str(row_number_1)).color = COLOR_CYAN
                #Building Data for Price Up, Volume Up
                price_vol_up_list.append(key)
            elif cum_price_diff_dict[key] > 0 and (cum_vol_diff_dict[key] - prev_cum_vol_diff_dict[key]) < 0:
                puvu.range(f'{get_col_name(col_number_2)}' + str(row_number_1)).value = "Up"
                puvu.range(f'{get_col_name(col_number_2)}' + str(row_number_1)).color = COLOR_GREEN
                puvu.range(f'{get_col_name(col_number_2+1)}' + str(row_number_1)).value = "Down"
                puvu.range(f'{get_col_name(col_number_2+1)}' + str(row_number_1)).color = COLOR_RED
                #Building Data for Price Up, Volume Down
                price_up_vol_down_list.append(key)
            else:
                logger.debug("Scenario not handled")
        col_number_2 += 2

    #Printing Stocks for which Price and Volume are Up/Down, along with subsequent filtering as per Support and Resistance.
    if row_number_1 > 2:
        puvu.range(f'B{temp_row_number + 1}').value = price_vol_up_list
        pdvd.range(f'B{row_number_1 - 1}').value = price_vol_down_list
        puvd.range(f'B{row_number_1 - 1}').value = price_up_vol_down_list
        pdvu.range(f'B{row_number_1 - 1}').value = price_down_vol_up_list
        temp_df = get_equity_config_df()
        temp_col_number = 2
        temp_col_number_1 = 2
        temp_col_number_2 = 2
        temp_col_number_3 = 2
        if temp_df is not None:
            #temp_list = []
            """if row_number_1 > 3 and price_vol_dir_supp_res_df is not None:
                validate_buy_sell(df, price_vol_dir_supp_res_df, price_vol_up_list, price_vol_down_list)
                price_vol_dir_supp_res_df = None
            price_vol_dir_supp_res_df = pd.DataFrame(index=temp_df.index, columns=['PUVU_RES','PUVU_SUP','PDVD_RES','PDVD_SUP'])"""
            for idx in temp_df.index:
                if idx in price_vol_up_list:
                    logger.debug(f'Configured Equity {idx} is in Price Up,Volume Up list')
                    puvu.range(f'{get_col_name(temp_col_number)}' + str(temp_row_number + row_difference + 1)).value = idx
                    if df.loc[idx, 'last_price'] >= temp_df.loc[idx, 'Resistance']:
                        logger.debug("Configured Equity price is greater than or equal to configured Resistance")
                        puvu.range(f'{get_col_name(temp_col_number)}' + str(temp_row_number + row_difference + 1)).color = COLOR_GREEN
                        #temp_list.append([df.loc[idx, 'last_price'],'NA','NA','NA'])
                        #price_vol_dir_supp_res_df.loc[idx] = [df.loc[idx, 'last_price'],'NA','NA','NA']
                    elif df.loc[idx, 'last_price'] <= temp_df.loc[idx, 'Support']:
                        logger.debug("Configured Equity price is less than or equal to configured Support")
                        puvu.range(f'{get_col_name(temp_col_number)}' + str(temp_row_number + row_difference + 1)).color = COLOR_RED
                        #price_vol_dir_supp_res_df.loc[idx] = ['NA',df.loc[idx, 'last_price'],'NA','NA']
                    else:
                        logger.debug("Configured Equity price is with-in Resistance and Support")
                        puvu.range(f'{get_col_name(temp_col_number)}' + str(temp_row_number + row_difference +1 )).color = COLOR_YELLOW
                        #price_vol_dir_supp_res_df.loc[idx] = ['NA','NA','NA','NA']
                    temp_col_number += 1
                elif idx in price_vol_down_list:
                    logger.debug(f'Configured Equity {idx} is in Price Down,Volume Down list')
                    pdvd.range(f'{get_col_name(temp_col_number_1)}' + str(temp_row_number + 1)).value = idx
                    if df.loc[idx, 'last_price'] >= temp_df.loc[idx, 'Resistance']:
                        logger.debug("Configured Equity price is greater than or equal to configured Resistance")
                        pdvd.range(f'{get_col_name(temp_col_number_1)}' + str(temp_row_number + 1)).color = COLOR_GREEN
                        #price_vol_dir_supp_res_df.loc[idx] = ['NA','NA',df.loc[idx, 'last_price'],'NA']
                    elif df.loc[idx, 'last_price'] <= temp_df.loc[idx, 'Support']:
                        logger.debug("Configured Equity price is less than or equal to configured Support")
                        pdvd.range(f'{get_col_name(temp_col_number_1)}' + str(temp_row_number + 1)).color = COLOR_RED
                        #price_vol_dir_supp_res_df.loc[idx] = ['NA','NA','NA',df.loc[idx, 'last_price']]
                    else:
                        logger.debug("Configured Equity price is with-in Resistance and Support")
                        pdvd.range(f'{get_col_name(temp_col_number_1)}' + str(temp_row_number + 1 )).color = COLOR_YELLOW
                        #price_vol_dir_supp_res_df.loc[idx] = ['NA','NA','NA','NA']
                    temp_col_number_1 += 1
                elif idx in price_up_vol_down_list:
                    logger.debug(f'Configured Equity {idx} is in Price Up,Volume Down list')
                    puvd.range(f'{get_col_name(temp_col_number_2)}' + str(temp_row_number + 1)).value = idx
                    temp_col_number_2 += 1
                elif idx in price_down_vol_up_list:
                    logger.debug(f'Configured Equity {idx} is in Price Down,Volume Up list')
                    pdvu.range(f'{get_col_name(temp_col_number_3)}' + str(temp_row_number + 1)).value = idx
                    temp_col_number_3 += 1
                else:
                    logger.debug(f'Configured Equity {idx} is not present in any list')
            #logger.debug(f'Price Volume Direction Support Resistance DataFrame - {price_vol_dir_supp_res_df}')
        else:
            logger.debug("Function create_price_vol_sheets, Equity Config df is None")

############################# End - Function to compare Cumulative Price, Volume difference after every fixed duration ############ 

############################# Start - Function to print top 5 gainers and loosers ############
def print_top_gainers_loosers(df):
    gainers_df = df.sort_values(by="percent_change", ascending = False)
    loosers_df = df.sort_values(by="percent_change")
    #logger.debug(f'top 5 gainsers - {gainers_df.head(5)}')
    #logger.debug(f'top 5 loosers - {loosers_df.head(5)}')
    eq.range("D31").value = "Top Gainers"
    eq.range('D31').font.bold = True
    eq.range("F31").value = "Top Loosers"
    eq.range('F31').font.bold = True
    eq.range("D32").value = gainers_df.head(5)['percent_change']    
    eq.range("F32").value = loosers_df.head(5)['percent_change']
############################# End - Function to print top 5 gainers and loosers ############
"""
############################# Start - Function to print equity info ############
def print_equity_info(df, json_data, eq_symbol):
    ret = True
	#bid_list = ask_list = trd_data = []
    bid_list = ask_list = []
    trd_data = []
    security_wise_dp = []
    tot_buy = tot_sell = 0
    eq.range("G3").value = df.loc[eq_symbol,'last_price']
    for key,value in json_data.items():
        if str(key) == "marketDeptOrderBook":
            for k,v in value.items():
                if str(k) == "bid":
                    bid_list = v
                elif str(k) == "ask":
                    ask_list = v
                elif str(k) == "tradeInfo":
                    trd_data.append(v)
                elif str(k) == "totalBuyQuantity":
                    tot_buy = v
                elif str(k) == "totalSellQuantity":
                    tot_sell = v
        elif str(key) == "securityWiseDP":
            security_wise_dp.append(value)
    
    bid_df = pd.DataFrame(bid_list)
    bid_df.rename(columns={"price":"Bid Price","quantity":"Bid Quantity"},inplace=True)
    ask_df = pd.DataFrame(ask_list)
    ask_df.rename(columns={"price":"Ask Price","quantity":"Ask Quantity"},inplace=True) 
    bid_ask_df = pd.concat([bid_df,ask_df], axis=1)
    trd_df = pd.DataFrame(trd_data)
    trd_df.drop(["marketLot", "activeSeries"], axis=1, inplace=True)
    trd_df = trd_df.transpose()
    security_wise_dp_df = pd.DataFrame(security_wise_dp)
    security_wise_dp_df.drop(["seriesRemarks"], axis=1, inplace=True)
    security_wise_dp_df = security_wise_dp_df.transpose()
    try:
        eq.range("D5").value = "Trade Data"
        eq.range('D5').font.bold = True
        eq.range("D6").value = trd_df
        eq.range("E6").value = None
        eq.range("F7").value = "Lakhs"
        eq.range("F8").value = "₹ Cr"
        eq.range("F9").value = "₹ Cr"
        eq.range("F10").value = "₹ Cr"
        eq.range("D15").value = "Delivery Data"
        eq.range('D15').font.bold = True
        eq.range("D16").value = security_wise_dp_df
        eq.range("E16").value = None
        eq.range("F19").value = "%"
        eq.range("D22").value = "Bid/Ask"
        eq.range('D22').font.bold = True
        eq.range("D23").options(pd.DataFrame, index=False).value = bid_ask_df
        eq.range("D29").value = "TotalBuyQty"
        eq.range("E29").value = tot_buy
        eq.range("F29").value = "TotalSellQty"
        eq.range("G29").value = tot_sell
        ret = True
    except Exception as e:
        logger.error(f'Error printing trading info df - {e}')
        ret = False
    
    return ret
############################# End - Function to print equity info ############ 
"""
while True:
    time.sleep(10)
    ############################# OptionChain Starts #############################
    """
    try:
        oc_sym, oc_exp = oc.range("E2").value, oc.range("E3").value
        OPTION_PCR_DURATION = cfg.range('B13').value
        RISK_FREE_INT_RATE = cfg.range('B14').value
        RISK_FREE_INT_RATE = RISK_FREE_INT_RATE/100
    except Exception as e:
        logger.debug(f'Closing Excel and handling exception - {e}')
        sys.exit()  
    if pre_oc_sym != oc_sym or pre_oc_exp != oc_exp:
        oc.range("H1:AD4000").value = None
        if pre_oc_sym != oc_sym:
            oc.range("B:B").value = oc.range("D6:E19").value = None
            exp_list = []
            oc_exp = None
        pre_oc_sym = oc_sym
        pre_oc_exp = oc_exp
    df = None  
    if oc_sym is not None:
        indices = True if oc_sym == "NIFTY" or oc_sym == "BANKNIFTY" else False        
        if not exp_list:
            logger.debug('Creating the Options expiry list')
            for i in list(nse.options_data(oc_sym, indices)["expiryDate"]):
                if dateutil.parser.parse(i).date() not in exp_list:
                    exp_list.append(dateutil.parser.parse(i).date())
            if exp_list:
                exp_df = pd.DataFrame({"Expiry Date": [str(i) for i in sorted(exp_list)]})
                exp_df = exp_df.set_index("Expiry Date", drop=True)
                oc.range("B1").value = exp_df
                oc.range("B1").autofit()
                logger.debug('Options expiry list created')
            else:
                logger.error(f'Error getting Options Expiry Dates - {e}')
                time.sleep(5)
                logger.debug("Trying to connect again...")
                nse = NSE()
                continue

        df = nse.options_data(oc_sym, indices)
        logger.debug(f'Expiry date input is - {oc_exp}')
        if df is not None and oc_exp is not None:
            logger.debug(f'DF is not none and Expiry date input is {oc_exp}')
            df["expiryDate"] = df["expiryDate"].apply(lambda x: dateutil.parser.parse(x))
            df = df[df["expiryDate"] == oc_exp]
            timestamp = list(df["timestamp"])[0]
            underlying_value = list(df["underlyingValue"])[0]

            ce_df = df[df["instrumentType"] == "CE"]
            ce_df = ce_df[["totalTradedVolume","change","lastPrice","impliedVolatility","changeinOpenInterest","openInterest","strikePrice"]]
            ce_df = ce_df.rename(columns={"openInterest":"CE OI", "changeinOpenInterest":"CE Change in OI", "impliedVolatility":"CE IV",
                                          "lastPrice":"CE LTP", "change":"CE LTP Change", "totalTradedVolume":"CE Volume"})
            ce_df.set_index("strikePrice", drop=True, inplace=True)
            ce_df["Strike"] = ce_df.index

            ce_df_greeks = get_option_greeks(ce_df, 'c', oc_exp)
            ce_df_greeks = ce_df_greeks.rename(columns={"Delta":"CE Delta", "Gamma":"CE Gamma", "Theta":"CE Theta", "Vega":"CE Vega",
                                                        "Rho":"CE Rho"})            
            ce_df_final = pd.concat([ce_df_greeks,ce_df], axis=1).sort_index()

            pe_df = df[df["instrumentType"] == "PE"]
            pe_df = pe_df[["strikePrice","openInterest","changeinOpenInterest","impliedVolatility","lastPrice","change","totalTradedVolume"]]
            pe_df = pe_df.rename(columns={"openInterest":"PE OI", "changeinOpenInterest":"PE Change in OI", "impliedVolatility":"PE IV",
                                          "lastPrice":"PE LTP", "change":"PE LTP Change", "totalTradedVolume":"PE Volume"})
            pe_df.set_index("strikePrice", drop=True, inplace=True)

            pe_df_greeks = get_option_greeks(pe_df, 'p', oc_exp)
            pe_df_greeks = pe_df_greeks[["Rho", "Vega", "Theta", "Gamma", "Delta"]]
            pe_df_greeks = pe_df_greeks.rename(columns={"Delta":"PE Delta", "Gamma":"PE Gamma", "Theta":"PE Theta", "Vega":"PE Vega",
                                                        "Rho":"PE Rho"})
            pe_df_final = pd.concat([pe_df, pe_df_greeks], axis=1).sort_index()

            df = pd.concat([ce_df_final,pe_df_final], axis=1).sort_index()
            df = df.replace(np.nan, 0)
            df["Strike"] = df.index
            df.index = [np.nan] * len(df)

            try:
                oc.range("D6").value = [["Spot LTP", underlying_value],
                                        ["Total Call OI", sum(list(df["CE OI"]))],
                                        ["Total Put OI", sum(list(df["PE OI"]))],
                                        ["",""],
                                        ["Max Call OI", max(list(df["CE OI"]))],
                                        ["Max Put OI", max(list(df["PE OI"]))],
                                        ["Max Call OI Strike", list(df[df["CE OI"] == max(list(df["CE OI"]))]["Strike"])[0]],
                                        ["Max Put OI Strike", list(df[df["PE OI"] == max(list(df["PE OI"]))]["Strike"])[0]],
                                        ["",""],
                                        ["Max Call Change in OI", max(list(df["CE Change in OI"]))],
                                        ["Max Put Change in OI", max(list(df["PE Change in OI"]))],
                                        ["Max Call Change in OI Strike",
                                        list(df[df["CE Change in OI"] == max(list(df["CE Change in OI"]))]["Strike"])[0]],
                                        ["Max Put Change in OI Strike",
                                        list(df[df["PE Change in OI"] == max(list(df["PE Change in OI"]))]["Strike"])[0]]
                                        ]
                oc.range("E1").value = timestamp
                option_curr_time = oc.range("E1").value
                oc.range("H1").value = df
                if option_prev_time != None and option_curr_time != None and option_prev_time != option_curr_time:
                    option_duration = option_curr_time - option_prev_time
                    #pcr_print_flag = False
                if option_row_number == 1 or (option_row_number > 1 and option_duration is not None and option_duration.total_seconds()/60 >= OPTION_PCR_DURATION):
                    oc.range(f'D{option_row_number + 22}').value = option_curr_time.strftime('%H:%M:%S')
                    oc.range(f'D{option_row_number + 22}').font.bold = True
                    pcr = sum(list(df["PE OI"]))/sum(list(df["CE OI"]))
                    oc.range(f'E{option_row_number + 22}').value = pcr
                    #pe_oi = get_option_oi(pe_df, 'p', underlying_value)
                    #ce_oi = get_option_oi(ce_df, 'c', underlying_value)
                    pe_oi = max(list(df["PE OI"]))
                    ce_oi = max(list(df["CE OI"]))
                    oc.range(f'F{option_row_number + 22}').value = pe_oi
                    oc.range(f'G{option_row_number + 22}').value = ce_oi
                    if prev_pcr is not None:
                        if pcr > prev_pcr:
                            oc.range(f'E{option_row_number + 22}').color = COLOR_GREEN
                        elif pcr < prev_pcr:
                            oc.range(f'E{option_row_number + 22}').color = COLOR_RED
                        else:
                            oc.range(f'E{option_row_number + 22}').color = COLOR_YELLOW
                    prev_pcr = pcr
                    option_prev_time = option_curr_time
                    option_row_number += 1
                    option_duration = None
            except Exception as e:
                logger.error(f'Error printing values - {e}')
                continue
        else:
            logger.error(f'Error getting Options Data - Either Options DataFrame is Null or Expiry date is not entered')
            if df is None:
                time.sleep(5)
                logger.debug("Trying to connect again...")
                nse = NSE()
                continue           
    ####################### OptionChain Ends ###########################
    """

    ####################### EquityData Starts ###########################
    try:
        ind_sym, eq_sym = eq.range("E2").value, eq.range("E3").value
        MAX_CUMULATIVE_TURNOVER = cfg.range('B3').value
        MAX_CUMULATIVE_TURNOVER = MAX_CUMULATIVE_TURNOVER*ONE_CRORE
        MIN_CUMULATIVE_TURNOVER = cfg.range('B4').value
        MIN_CUMULATIVE_TURNOVER = MIN_CUMULATIVE_TURNOVER*ONE_CRORE
        DURATION = cfg.range('B5').value
        PERCENT_UP = cfg.range('F3').value
        PERCENT_DOWN = cfg.range('F4').value
    except Exception as e:
        logger.debug(f'Closing Excel and handling exception - {e}')
        sys.exit()
    
    """
    if pre_ind_sym is not None and pre_ind_sym != ind_sym:
        eq_sym = None
        eq.range("I1:AE510").value = eq.range("D5:H39").value = None
        eq.range("E1").value = eq.range("G1").value = None
        eq.range("E3").value = eq.range("G2").value = None
        sv.clear()
        sp.clear()
        st.clear()
        maxct.clear()
        minct.clear()
        puvu.clear()
        row_number = 1
        col_number = 2
        row_number_1 = 2
        col_number_1 = 1
        prev_time = curr_time = None
        prev_time_1 = datetime.now()
        prev_time_2 = None
        prev_vol = curr_vol = []
        prev_vol_diff = curr_vol_diff = []
        prev_price = curr_price = []
        prev_price_diff = curr_price_diff = []
        prev_turn = curr_turn = []
        prev_turn_diff = curr_turn_diff = []
        cum_turn_dict = {}
        cum_price_diff_dict = {}
        prev_cum_price_diff_dict = {}
        cum_vol_diff_dict = {}
        prev_cum_vol_diff_dict = {}
        price_vol_dict_flag = False

    if pre_eq_sym is not None and pre_eq_sym != eq_sym:
        eq.range("D5:H39").value = None
        eq.range("G3").value = None
    pre_ind_sym = ind_sym
    pre_eq_sym = eq_sym

    """
    
    eq_df = None
    if ind_sym is not None:
        #eq_df = nse.equity_market_data(ind_sym)
        eq_df = kite.get_nifty50_market_data()
        if eq_df is not None:
            #eq_df.drop(["priority","date365dAgo","chart365dPath","date30dAgo","chart30dPath","chartTodayPath","series","identifier"],
                       #axis=1,inplace=True)
            #eq_df.index.name = 'symbol'
            #sorted_idx = eq_df.index.sort_values()
            #eq_df = eq_df.loc[sorted_idx]
            rows_eq_df = len(eq_df.index)
            #duration_1 = None
            try:
                eq.range("I1").value = eq_df
                eq.range("E1").value = eq_df.loc[ind_sym,'timestamp']
                eq.range("G2").value = eq_df.loc[ind_sym,'last_price']
                eq.range("G1").value = eq_df.iloc[0]['timestamp']
                curr_time = eq.range("G1").value
                if curr_time is None:
                    curr_time = eq.range("E1").value
                print_top_gainers_loosers(eq_df)
                
                """
                if prev_time_2 != None and curr_time != None and prev_time_2 != curr_time:
                    duration_1 = curr_time - prev_time_2
                if row_number == 1 or (row_number > 1 and duration_1 is not None and duration_1.total_seconds()/60 >= DELIVERY_CHANGE_DURATION):
                    eq.range("AB1").options(index=False).value = get_delivery_info(eq_df)
                    prev_time_2 = curr_time
                """
            except Exception as e:
                logger.error(f'Error printing df values - {e}')
                continue
            
            """
            data = None
            if eq_sym is not None:                
                data = nse.equity_info(eq_sym, trade_info=True)
                if data is not None:
                    ret_val = print_equity_info(eq_df, data, eq_sym)
                    if not ret_val:
                        continue
                else:
                    logger.error(f'Error getting Equity Info for {eq_sym} - Equity Info Data is Null')
                    time.sleep(5)
                    logger.debug("Trying to connect again...")
                    nse = NSE()
                    continue
            """
                    
            if row_number == 1:
                initial_rows_eq_df = len(eq_df.index)
            logger.debug(f'For Time {curr_time} and row number {row_number}, initial stocks - {initial_rows_eq_df} and current stocks - {rows_eq_df}')

            ####################### Start - Spot Data (Price,Volume,Turnover) ###########################
            if prev_time != curr_time and initial_rows_eq_df == rows_eq_df:
                stock_list = []
                duration = curr_time - prev_time_1
                try:
                    create_spot_data(eq_df,"Price",curr_time,duration,row_number,prev_price,curr_price,prev_price_diff,curr_price_diff,col_number_1,stock_list)                               
                    create_spot_data(eq_df,"Volume",curr_time,duration,row_number,prev_vol,curr_vol,prev_vol_diff,curr_vol_diff,col_number_1,stock_list)
                    create_spot_data(eq_df,"Turnover",curr_time,duration,row_number,prev_turn,curr_turn,prev_turn_diff,curr_turn_diff,col_number_1,stock_list)
                except Exception as e:
                    logger.error(f'Error Creating Spot data - {e}')
                    continue
                if row_number >= 2:
                    #col_number_1 += 4
                    price_vol_dict_flag = True
                
                if duration.total_seconds()/60 >= DURATION:
                    col_number_1 += 3
                    prev_time_1 = curr_time
                    cum_turn_dict = {}
                    create_price_vol_sheets(eq_df, curr_time, row_number_1)
                    prev_cum_price_diff_dict = cum_price_diff_dict
                    cum_price_diff_dict = {}
                    prev_cum_vol_diff_dict = cum_vol_diff_dict
                    cum_vol_diff_dict = {}
                    row_number_1 += 1
                    price_vol_dict_flag = False
                    prev_eq_df = eq_df

                if row_number > 1:
                    prev_price = curr_price
                    prev_vol = curr_vol
                    prev_turn = curr_turn
                curr_price = []
                curr_vol = []
                curr_turn = []

                if row_number > 2:
                    prev_price_diff = curr_price_diff
                    prev_vol_diff = curr_vol_diff
                    prev_turn_diff = curr_turn_diff
                curr_price_diff = []
                curr_vol_diff = []
                curr_turn_diff = []
                row_number += 1                    
            prev_time = curr_time
        else:
            logger.error(f'Error getting Equity Data - Equity DataFrame is Null')
            time.sleep(5)
            logger.debug("Trying to connect again...")
            kite = KiteMain()
            continue
            ####################### End - Spot Data (Price,Volume,Turnover) ###########################               
    ####################### EquityData Ends ###########################
"""
    ####################### FuturesData Starts ###########################
    try:
        fd_sym = fd.range("E2").value
    except Exception as e:
        logger.debug(f'Closing Excel and handling exception - {e}')
        sys.exit()
    if pre_fd_sym != fd_sym:
        fd.range("G1:AD100").value = None
        pre_fd_sym = fd_sym
    deriv_data = None
    if fd_sym is not None:
        indices = True if fd_sym == "NIFTY" or fd_sym == "BANKNIFTY" else False
        deriv_data = nse.derivatives_data(fd_sym)
        if deriv_data is not None:
            meta_data_list = []
            trd_info_list = []
            for i in deriv_data["stocks"]:
                if i["metadata"]["instrumentType"] == ("Index Futures" if indices else "Stock Futures"):
                    meta_data_list.append(i["metadata"])
                    trd_info_list.append(i["marketDeptOrderBook"]["tradeInfo"])

            meta_data_df = pd.DataFrame(meta_data_list)
            trd_info_df = pd.DataFrame(trd_info_list)
            meta_data_df = meta_data_df.set_index("identifier", drop=True)
            meta_data_df.drop(["optionType","strikePrice","closePrice"],axis=1,inplace=True)
            trd_info_df.drop(["tradedVolume","value","premiumTurnover","marketLot"],axis=1,inplace=True)
            try:
                fd.range("I1").value = meta_data_df
                fd.range("U1").options(index=False).value = trd_info_df
                deriv_timestamp = deriv_data["fut_timestamp"]
                fd.range("E1").value = deriv_timestamp
                deriv_underlying_value = deriv_data["underlyingValue"]
                fd.range("G1").value = deriv_underlying_value
            except Exception as e:
                logger.error(f'Error printing futures df - {e}')
                continue
        else:
            logger.error(f'Error getting Futures Data - Futures DataFrame is Null')
            time.sleep(5)
            logger.debug("Trying to connect again...")
            nse = NSE()
            continue
    ####################### FuturesData Ends ###########################

    """