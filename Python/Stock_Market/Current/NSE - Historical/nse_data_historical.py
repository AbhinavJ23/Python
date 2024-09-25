from NseIndia import NSE
import os,sys,time
import pandas as pd
import xlwings as xw
import dateutil.parser
import numpy as np
import time
#import logging
from datetime import datetime, timedelta
from base_logger import logger
import ctypes
#from py_vollib.black_scholes.implied_volatility import implied_volatility
from py_vollib.black_scholes.greeks.analytical import delta,gamma,rho,theta,vega

####################### Initializing Logging Start #######################
#logging.basicConfig(filename='Nse_Data_Historical_'+time.strftime('%Y%m%d%H%M%S')+'.log', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
#logger = logging.getLogger()
####################### Initializing Logging End #######################

############################# Start - Function to check validity,expiry #############################
def check_validity():
    valid_from_str = '23/09/2024 00:00:00'
    valid_from_time = datetime.strptime(valid_from_str, '%d/%m/%Y %H:%M:%S')
    #valid_from_time = datetime(2024, 8, 15, 0, 0, 0)
    #duration = timedelta(days=5, hours=0, minutes=0, seconds=0)
    #valid_till_str = '17/08/2024 22:30:30'   
    valid_till_time = valid_from_time + timedelta(days=7)
    time_now = datetime.now()
    time_left = valid_till_time - time_now
    logger.debug(f'Time Left - {time_left}')
    logger.debug(f'Days left - {time_left.days}')
    total_seconds = time_left.total_seconds()
    logger.debug(f'Total Seconds Left - {total_seconds}')
    if total_seconds < 0:
        ctypes.windll.user32.MessageBoxW(0, "Your product trial period has expired!", "Error",0)
        return False
    else:    
        #hours = round((total_seconds - time_left.days*24*60*60)/3600)
        #minutes = round((total_seconds // 60) % 60)
        #seconds = round(total_seconds % 60)
        hours = int((total_seconds - time_left.days*24*60*60)//3600)
        minutes = int((total_seconds - time_left.days*24*60*60 - hours*60*60)//60)
        seconds = round(total_seconds - time_left.days*24*60*60 - hours*60*60 - minutes*60)
        message = "Your product trial period will expire in " + str(time_left.days) + " day(s) " + str(hours) +" hours(s) " + str(minutes) + " min(s) and " + str(seconds) + " second(s)"
        ctypes.windll.user32.MessageBoxW(0, message, "Warning",0)
        return True
############################# End - Function to check validity,expiry #############################

status = check_validity()
if not status:
   sys.exit()

nse = NSE()

## Creating new excel and adding sheets
file_name = 'Nse_Data_Historical_'+time.strftime('%Y%m%d%H%M%S')+'.xlsx'
if not os.path.exists(file_name):
    try:
        wb = xw.Book()
        wb.sheets.add("FuturesData")
        wb.sheets.add("OptionChain")
        wb.sheets.add("EquityData")
        wb.save(file_name)
        #wb.close()
        logger.debug("Created Excel - " + file_name)
    except Exception as e:
        logger.critical(f'Error Creating Excel - {e}')
        sys.exit()

wb = xw.Book(file_name)
eq = wb.sheets("EquityData")
oc = wb.sheets("OptionChain")
fd = wb.sheets("FuturesData")

####################### Initializing Constants #######################
COLOR_GREY = (211, 211, 211)
COLOR_GREEN = (0, 255, 0)
COLOR_RED = (255, 0, 0)
COLOR_YELLOW = (255, 255, 0)
RISK_FREE_INT_RATE = 0.05
DELIVERY_CHANGE_DURATION = 60

####################### Initializing Excel Sheets #######################
oc.range('1:1').font.bold = True
oc.range('1:1').color = COLOR_GREY
oc.range('C1:C200').color = COLOR_GREY
oc.range('G1:G50000').color = COLOR_GREY
oc.range('C1').column_width = 2
oc.range('G1').column_width = 2
eq.range('1:1').font.bold = True
eq.range('1:1').color = COLOR_GREY
eq.range('B1:C34').color = COLOR_GREY
eq.range('B1').column_width = 1
eq.range('C1').column_width = 1
eq.range('H1').column_width = 2
eq.range('H1:H40000').color = COLOR_GREY
fd.range('1:1').font.bold = True
fd.range('1:1').color = COLOR_GREY
fd.range('B1:C200').color = COLOR_GREY
fd.range('B1').column_width = 1
fd.range('C1').column_width = 1
fd.range('H1').column_width = 2
fd.range('H1:H1000').color = COLOR_GREY

####################### Initializing OptionChain sheet #######################
oc.range("A:B").value = oc.range("D6:E19").value = oc.range("G1:V4000").value = None
oc_df = None
#try:    
oc_df= pd.DataFrame({"FNO Symbol":["NIFTY", "BANKNIFTY"] + nse.equity_market_data("Securities in F&O", symbol_list=True)})
#except Exception as e:
    #logger.critical(f'Error getting FNO symbols for Options Data - {e}')
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
oc.range('D20:F20').color = COLOR_GREY
oc.range("D1").value = "Current Time"
pre_oc_sym = pre_oc_exp = None
exp_list = []
logger.debug("OptionChain sheet initialized")

######################### Initializing EquityData sheet #######################
eq.range("A:A").value = eq.range("D5:H30").value = eq.range("I1:AD510").value = None
eq_df = pd.DataFrame({"Index Symbol":nse.equity_market_categories})
eq_df = eq_df.set_index("Index Symbol", drop=True)
eq.range("A1").value = eq_df
eq.range("A1:A50").autofit()
eq.range("D2").value, eq.range("D3").value = "Enter Index ->", "Enter Equity ->"
eq.range('D2').font.bold = True
eq.range('D3').font.bold = True
eq.range('A35:G35').color = COLOR_GREY
eq.range("D1").value = "Index Time"
eq.range("D1").autofit()
eq.range("F1").value = "Equity Time"
eq.range("F1").autofit()
eq.range("F2").value = "Index Value"
eq.range('F2').font.bold = True
pre_ind_sym = pre_eq_sym = None
eq.range("F3").value = "Equity Value"
eq.range('F3').font.bold = True
eq.range("D2:E3").autofit()
logger.debug("EquityData sheet initialized")

####################### Initializing FuturesData sheet #######################
fd.range("A:A").value = fd.range("G1:AD100").value = None
fd_df = None
#try:
fd_df= pd.DataFrame({"FNO Symbol":["NIFTY", "BANKNIFTY"] + nse.equity_market_data("Securities in F&O", symbol_list=True)})
#except Exception as e:
    #logger.critical(f'Error getting FNO symbols for Futures Data - {e}')
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
logger.debug("All Excel sheets initialized")

####################### Initializing Global Variables #######################
eq_row_number = 1
oc_row_number = 1
fd_row_number = 1
eq_prev_time = eq_curr_time = None
eq_prev_time_1 = None
oc_prev_time = oc_curr_time = None
fd_prev_time = fd_curr_time = None
eq_df_flag = True
oc_df_flag = True
fd_df_flag = True

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

############################# Start - Function to get option greeks #############################
def get_option_greeks(df, call_or_put, expiry):
    time = ((datetime(expiry.year, expiry.month, expiry.day, 15, 30) - datetime.now()) / timedelta(days=1)) / 365
    #time = (datetime(expiry.year, expiry.month, expiry.day, 15, 30) - datetime.now()).total_seconds() / (60*60*24*365)
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
            #logger.debug(f'underlying_price - {last_price}, Strike - {strike}, time - {time}, imp vol - {imp_vol}')
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

############################# Start - Function to get delivery info ###########################
def get_delivery_info(df):
    delivery_info_list = []
    flag = True
    for i in df.index:
        symbol = i
        data = None
        if symbol != 'NIFTY 50':
            data = nse.equity_info(symbol, trade_info=True)
            if data is not None:
                for key,value in data.items():              
                    if str(key) == "securityWiseDP":
                        delivery_info_list.append(value)
            else:
                logger.error(f'Error getting Delivery Info for {symbol} - Delivery Info Data is Null')
                flag = False
                break
                #delivery_info_list.append({"quantityTraded": 'Error', "deliveryQuantity": 'Error', "deliveryToTradedQuantity": 'Error'})
                #time.sleep(5)
                #logger.debug("Trying to connect again...")
                #nse = NSE()
                #continue  
                #empty_df = pd.DataFrame(index=df.index, columns=['quantityTraded','deliveryQuantity','deliveryToTradedQuantity'])
                #empty_df.fillna(0)
                #return empty_df
        else:
            delivery_info_list.append({"quantityTraded": 'NA', "deliveryQuantity": 'NA', "deliveryToTradedQuantity": 'NA'})
    
    if len(delivery_info_list) == len(df.index) and flag:
        delivery_info_df = pd.DataFrame(delivery_info_list, index=df.index)
        delivery_info_df.drop(["seriesRemarks","secWiseDelPosDate"],axis=1,inplace=True)
        return delivery_info_df
    else:
        empty_df = pd.DataFrame(index=df.index, columns=['quantityTraded','deliveryQuantity','deliveryToTradedQuantity'])
        empty_df.fillna(0)
        return empty_df

############################# End - Function to get delivery info #############################

while True:
    time.sleep(1)
    ############################# OptionChain Starts #############################
    try:
        oc_sym, oc_exp = oc.range("E2").value, oc.range("E3").value
    except Exception as e:
        logger.debug(f'Closing Excel and handling exception - {e}')
        sys.exit()
    if pre_oc_sym != oc_sym or pre_oc_exp != oc_exp:
        oc.range("G1:AD50000").value = None
        oc_row_number = 1
        oc_prev_time = oc_curr_time = None
        oc_df_flag = True
        if pre_oc_sym != oc_sym:
            oc.range("B:B").value = oc.range("D6:E19").value = None
            exp_list = []
            oc_exp = None
        pre_oc_sym = oc_sym
        pre_oc_exp = oc_exp
    df = None    
    if oc_sym is not None:
        indices = True if oc_sym == "NIFTY" or oc_sym == "BANKNIFTY" else False
        logger.debug(f'value of Options row number is {oc_row_number}')
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
                #logger.debug(exp_df)
            else:
                logger.error(f'Error getting Options Expiry Dates - {e}')
                time.sleep(5)
                logger.debug("Trying to connect again...")
                nse = NSE()
                continue
        #try:
        #    logger.debug('Getting Options data')
        #    df = nse.options_data(oc_sym, indices)
        #except Exception as e:
        #    logger.error(f'Error getting Options Data - {e}')
        #    time.sleep(5)
        #    continue
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
            #logger.debug("CE DF Greeks:")
            #logger.debug(ce_df_greeks)
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
            #logger.debug("PE DF Greeks:")
            #logger.debug(pe_df_greeks)
            pe_df_final = pd.concat([pe_df, pe_df_greeks], axis=1).sort_index()

            df = pd.concat([ce_df_final,pe_df_final], axis=1).sort_index()
            df = df.replace(np.nan, 0)
            df["Strike"] = df.index
            df.index = [np.nan] * len(df)
            rows_oc_df = len(df.index)
            #oc.range("D6").value = [["Timestamp", timestamp],
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
                #oc.range("D1").value = "Current Time"
                oc.range("E1").value = timestamp
                oc_curr_time = oc.range("E1").value
                #oc.range("E1").autofit()           
                #oc.range("G1").value = df
            except Exception as e:
                logger.error(f'Error printing values - {e}')
                continue           
            if oc_row_number == 1 and oc_df_flag:
                logger.debug(f'Printing the options chain df for the first time at {oc_curr_time}')
                try:
                    oc.range("F1").value = timestamp
                    oc.range(f'G{oc_row_number}').value = df
                except Exception as e:
                    logger.error(f'Error printing options chain df - {e}')
                    continue
                oc_df_flag = False

            oc_duration = None
            if oc_prev_time != None and oc_curr_time != None and oc_prev_time != oc_curr_time:
                oc_duration = oc_curr_time - oc_prev_time

            logger.debug(f'Options : Previous time - {oc_prev_time}' + f',Current time - {oc_curr_time}' + f',Duration -  {oc_duration}')            
            #if oc_prev_time != None and oc_prev_time != oc_curr_time:                
            if oc_duration is not None and oc_duration.total_seconds() > 0:
                logger.debug(f'Printing the options chain df for the next time at {oc_curr_time}')
                oc_row_number += rows_oc_df
                oc_row_number += 1
                try:
                    oc.range(f'G{oc_row_number}' + ':' + f'AD{oc_row_number}').color = COLOR_GREY               
                    oc.range(f'G{oc_row_number}').value = df                            
                    oc.range(f'F{oc_row_number}').value = oc_curr_time
                    #oc.range(f'F{oc_row_number}').autofit()
                    oc.range(f'F{oc_row_number}').font.bold = True
                    oc.range(f'G{oc_row_number}' + ':' + f'AD{oc_row_number}').font.bold = True
                except Exception as e:
                    logger.error(f'Error printing options chain df - {e}')
                    continue

            if oc_row_number == 1 or (oc_row_number > 1 and oc_duration is not None and oc_duration.total_seconds() > 0):
                oc_prev_time = oc_curr_time
        else:
            logger.error(f'Error getting Options Data - Either Options DataFrame is Null or Expiry date is not entered')
            if df is None:
                time.sleep(5)
                logger.debug("Trying to connect again...")
                nse = NSE()
                continue            
    ####################### OptionChain Ends ###########################

    ####################### EquityData Starts ###########################
    try:
        ind_sym, eq_sym = eq.range("E2").value, eq.range("E3").value
    except Exception as e:
        logger.debug(f'Closing Excel and handling exception - {e}')
        sys.exit()
    if pre_ind_sym != ind_sym:
        eq_sym = None
        eq.range("I1:AD40000").value = eq.range("D5:H30").value = None
        eq.range("E1").value = eq.range("G1").value = None
        eq.range("E3").value = eq.range("G2").value = None
        eq_row_number = 1
        eq_df_flag = True
        eq_prev_time = eq_curr_time = None
        eq_prev_time_1 = None

    if pre_eq_sym != eq_sym:
        eq.range("D5:H30").value = None
        #eq.range("F3").value = None
        eq.range("G3").value = None
    pre_ind_sym = ind_sym
    pre_eq_sym = eq_sym
    eq_df = None 
    if ind_sym is not None:
        logger.debug(f'value of Equity row number is {eq_row_number}')
        #try:            
        #    eq_df = nse.equity_market_data(ind_sym)
        #except Exception as e:
        #    logger.error(f'Error getting Equity Data - {e}')
        #    time.sleep(5)
        #    continue
        eq_df = nse.equity_market_data(ind_sym)
        if eq_df is not None:            
            eq_df.drop(["priority","date365dAgo","chart365dPath","date30dAgo","chart30dPath","chartTodayPath","series","identifier"],
                       axis=1,inplace=True)
            eq_df.index.name = 'symbol'            
            sorted_idx = eq_df.index.sort_values()
            eq_df = eq_df.loc[sorted_idx]
            rows_eq_df = len(eq_df.index)
            #eq.range("I1").value = eq_df
            try:
                eq.range("E1").value = eq_df.loc[ind_sym,'lastUpdateTime']
                eq.range("G2").value = eq_df.loc[ind_sym,'lastPrice']            
                eq.range("G1").value = eq_df.iloc[0]['lastUpdateTime']
                eq_curr_time = eq.range("G1").value            
                #eq.range("G1").autofit()
            except Exception as e:
                logger.error(f'Error printing values - {e}')
                continue

            eq_duration = None
            if eq_prev_time_1 != None and eq_curr_time != None and eq_prev_time_1 != eq_curr_time:
                eq_duration = eq_curr_time - eq_prev_time_1

            if eq_row_number == 1 and eq_df_flag:
                logger.debug(f'Printing the Equity df for the first time at {eq_curr_time}')
                try:
                    eq.range(f'I{eq_row_number}').value = eq_df
                    eq.range(f'AA{eq_row_number}').options(index=False).value = get_delivery_info(eq_df)
                except Exception as e:
                    logger.error(f'Error printing equity df - {e}')
                    continue
                eq_df_flag = False

            if eq_prev_time != None and eq_prev_time != eq_curr_time:
                logger.debug(f'Printing the Equity df for the next time at {eq_curr_time}')
                eq_row_number += rows_eq_df
                eq_row_number += 1
                try:
                    eq.range(f'I{eq_row_number}' + ':' + f'AD{eq_row_number}').color = COLOR_GREY               
                    eq.range(f'I{eq_row_number}').value = eq_df                            
                    eq.range(f'G{eq_row_number}').value = eq_curr_time
                    eq.range(f'G{eq_row_number}').font.bold = True
                    eq.range(f'I{eq_row_number}' + ':' + f'AD{eq_row_number}').font.bold = True
                except Exception as e:
                    logger.error(f'Error printing equity df - {e}')
                    continue        
            if eq_duration is not None and eq_duration.total_seconds()/60 > DELIVERY_CHANGE_DURATION:
                try:
                    eq.range(f'AA{eq_row_number}').options(index=False).value = get_delivery_info(eq_df)
                except Exception as e:
                    logger.error(f'Error printing delivery info df - {e}')
                    continue

            if eq_row_number == 1 or (eq_row_number > 1 and eq_duration is not None and eq_duration.total_seconds()/60 > DELIVERY_CHANGE_DURATION):
                eq_prev_time_1 = eq_curr_time
            eq_prev_time = eq_curr_time

            data = None
            if eq_sym is not None:                
                #try:
                #    data = nse.equity_info(eq_sym, trade_info=True)
                #except Exception as e:
                #    logger.error(f'Error getting Equity Info for {eq_sym} - {e}')
                #    time.sleep(5)
                #    continue
                data = nse.equity_info(eq_sym, trade_info=True)
                if data is not None:
                    bid_list = ask_list = [] #trd_data = security_wise_dp = []
                    trd_data = []
                    security_wise_dp = []
                    tot_buy = tot_sell = 0
                    eq.range("G3").value = eq_df.loc[eq_sym,'lastPrice']
                    for key,value in data.items():
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
                    trd_df = pd.DataFrame(trd_data).transpose()
                    security_wise_dp_df = pd.DataFrame(security_wise_dp).transpose()
                    try:
                        eq.range("D5").value = trd_df
                        eq.range("E5").value = None
                        eq.range("F6").value = "Lakhs"
                        eq.range("F7").value = "₹ Cr"
                        eq.range("F8").value = "₹ Cr"
                        eq.range("F9").value = "₹ Cr"
                        eq.range("D15").value = security_wise_dp_df
                        eq.range("E15").value = None
                        eq.range("F18").value = "%"               
                        eq.range("D22").options(pd.DataFrame, index=False).value = bid_ask_df
                        eq.range("D28").value = "TotalBuyQty"
                        eq.range("E28").value = tot_buy
                        eq.range("F28").value = "TotalSellQty"
                        eq.range("G28").value = tot_sell
                    except Exception as e:
                        logger.error(f'Error printing trading info df - {e}')
                        continue                 
                else:
                    logger.error(f'Error getting Equity Info for {eq_sym} - Equity Info Data is Null')
                    time.sleep(5)
                    logger.debug("Trying to connect again...")
                    nse = NSE()
                    continue          
        else:
            logger.error(f'Error getting Equity Data - Equity DataFrame is Null')
            time.sleep(5)
            logger.debug("Trying to connect again...")
            nse = NSE()
            continue
    ####################### EquityData Ends ###########################

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
        #fd_df = nse.futures_data(fd_sym, indices)
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
            rows_fd_df = len(meta_data_df.index)
            #fd.range("I1").value = meta_data_df
            trd_info_df.drop(["tradedVolume","value","premiumTurnover","marketLot"],axis=1,inplace=True)
            #fd.range("U1").options(index=False).value = trd_info_df
            deriv_timestamp = deriv_data["fut_timestamp"]
            try:
                fd.range("E1").value = deriv_timestamp
                fd_curr_time = fd.range("E1").value
                deriv_underlying_value = deriv_data["underlyingValue"]
                fd.range("G1").value = deriv_underlying_value
            except Exception as e:
                    logger.error(f'Error printing values - {e}')
                    continue

            if fd_row_number == 1 and fd_df_flag:
                logger.debug(f'Printing the futures df for the first time at {fd_curr_time}')
                try:
                    fd.range(f'I{fd_row_number}').value = meta_data_df
                    fd.range(f'U{fd_row_number}').options(index=False).value = trd_info_df
                except Exception as e:
                    logger.error(f'Error printing futures df - {e}')
                    continue
                fd_df_flag = False

            fd_duration = None
            if fd_prev_time != None and fd_curr_time != None and fd_prev_time != fd_curr_time:
                fd_duration = fd_curr_time - fd_prev_time

            logger.debug(f'Futures : Previous time - {fd_prev_time}' + f',Current time - {fd_curr_time}' + f',Duration -  {fd_duration}')        
            if fd_duration is not None and fd_duration.total_seconds() > 0:
                logger.debug(f'Printing the futures df for the next time at {fd_curr_time}')
                fd_row_number += rows_fd_df
                fd_row_number += 1
                try:
                    fd.range(f'I{fd_row_number}' + ':' + f'X{fd_row_number}').color = COLOR_GREY               
                    fd.range(f'I{fd_row_number}').value = meta_data_df
                    fd.range(f'U{fd_row_number}').options(index=False).value = trd_info_df                            
                    fd.range(f'G{fd_row_number}').value = fd_curr_time
                    fd.range(f'G{fd_row_number}').font.bold = True
                    fd.range(f'I{fd_row_number}' + ':' + f'X{fd_row_number}').font.bold = True
                except Exception as e:
                    logger.error(f'Error printing futures df - {e}')
                    continue

            if fd_row_number == 1 or (fd_row_number > 1 and fd_duration is not None and fd_duration.total_seconds() > 0):
                fd_prev_time = fd_curr_time
        else:
            logger.error(f'Error getting Futures Data - Futures DataFrame is Null')
            time.sleep(5)
            logger.debug("Trying to connect again...")
            nse = NSE()
            continue
    ####################### FuturesData Ends ###########################




