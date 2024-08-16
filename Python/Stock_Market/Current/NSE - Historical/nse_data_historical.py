from NseIndia import NSE
import os,sys,time
import pandas as pd
import xlwings as xw
import dateutil.parser
import numpy as np
import time
import logging
from datetime import datetime, timedelta
from base_logger import logger
import ctypes

####################### Initializing Logging Start #######################
#logging.basicConfig(filename='Nse_Data_Historical_'+time.strftime('%Y%m%d%H%M%S')+'.log', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
#logger = logging.getLogger()
####################### Initializing Logging End #######################

############################# Start - Function to check validity,expiry #############################
def check_validity():
    valid_from_str = '15/08/2024 00:00:00'
    valid_from_time = datetime.strptime(valid_from_str, '%d/%m/%Y %H:%M:%S')
    #valid_from_time = datetime(2024, 8, 15, 0, 0, 0)
    #duration = timedelta(days=5, hours=0, minutes=0, seconds=0)
    #valid_till_str = '17/08/2024 22:30:30'   
    valid_till_time = valid_from_time + timedelta(days=5)
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
logger.debug("Excel sheets initialized")

####################### Initializing OptionChain sheet #######################
oc.range("A:B").value = oc.range("D6:E19").value = oc.range("G1:V4000").value = None
try:    
    oc_df= pd.DataFrame({"FNO Symbol":["NIFTY", "BANKNIFTY"] + nse.equity_market_data("Securities in F&O", symbol_list=True)})
except Exception as e:
    logger.critical(f'Error getting FNO symbols for Options Data - {e}')
oc_df = oc_df.set_index("FNO Symbol", drop=True)
oc.range("A1").value = oc_df
oc.range("A1:A200").autofit()
oc.range("D2").value, oc.range("D3").value = "Enter Symbol ->", "Enter Expiry ->"
oc.range('D2').font.bold = True
oc.range('D3').font.bold = True
oc.range("D2:E3").autofit()
oc.range('A200:B200').color = COLOR_GREY
oc.range('D20:F20').color = COLOR_GREY
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
try:
    fd_df= pd.DataFrame({"FNO Symbol":["NIFTY", "BANKNIFTY"] + nse.equity_market_data("Securities in F&O", symbol_list=True)})
except Exception as e:
    logger.critical(f'Error getting FNO symbols for Futures Data - {e}')
fd_df = fd_df.set_index("FNO Symbol", drop=True)
fd.range("A1").value = fd_df
fd.range("A1:A200").autofit()
fd.range("D2").value = "Enter Index/Equity ->"
fd.range('D2').font.bold = True
fd.range("D2").autofit()
pre_fd_sym = None
logger.debug("FuturesData sheet initialized")

####################### Initializing Global Variables #######################
eq_row_number = 1
oc_row_number = 1
eq_prev_time = eq_curr_time = None
oc_prev_time = oc_curr_time = None
eq_df_flag = True
oc_df_flag = True
#prev_time_1 = datetime.now()

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

while True:
    time.sleep(1)
    ############################# OptionChain Starts #############################
    oc_sym, oc_exp = oc.range("E2").value, oc.range("E3").value    
    if pre_oc_sym != oc_sym or pre_oc_exp != oc_exp:
        oc.range("G1:V4000").value = None
        if pre_oc_sym != oc_sym:
            oc.range("B:B").value = oc.range("D6:E19").value = None
            exp_list = []
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
                exp_df = pd.DataFrame({"Expiry Date": [str(i) for i in sorted(exp_list)]})
                exp_df = exp_df.set_index("Expiry Date", drop=True)
            if exp_list:
                oc.range("B1").value = exp_df
                oc.range("B1").autofit()
                logger.debug('Options expiry list created')
                logger.debug(exp_df)
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

            pe_df = df[df["instrumentType"] == "PE"]
            pe_df = pe_df[["strikePrice","openInterest","changeinOpenInterest","impliedVolatility","lastPrice","change","totalTradedVolume"]]
            pe_df = pe_df.rename(columns={"openInterest":"PE OI", "changeinOpenInterest":"PE Change in OI", "impliedVolatility":"PE IV",
                                          "lastPrice":"PE LTP", "change":"PE LTP Change", "totalTradedVolume":"PE Volume"})
            pe_df.set_index("strikePrice", drop=True, inplace=True)

            df = pd.concat([ce_df,pe_df], axis=1).sort_index()
            df = df.replace(np.nan, 0)
            df["Strike"] = df.index
            df.index = [np.nan] * len(df)
            rows_oc_df = len(df.index)
            #oc.range("D6").value = [["Timestamp", timestamp],
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
            oc.range("D1").value = "Timestamp"
            oc.range("E1").value = timestamp
            oc_curr_time = oc.range("E1").value
            #oc.range("E1").autofit()           
            #oc.range("G1").value = df
            if oc_row_number == 1 and oc_df_flag:
                logger.debug(f'Printing the options expiry df for the first time at {oc_curr_time}')
                oc.range(f'G{oc_row_number}').value = df
                oc_df_flag = False

            duration = None
            if oc_prev_time != None and oc_curr_time != None and oc_prev_time != oc_curr_time:
                duration = oc_curr_time - oc_prev_time

            logger.debug(f'Previous time - {oc_prev_time}' + f',Current time - {oc_curr_time}' + f',Duration -  {duration}')            
            #if oc_prev_time != None and oc_prev_time != oc_curr_time:                
            if duration is not None and duration.total_seconds() > 0:
                logger.debug(f'Printing the options expiry df for the next time at {oc_curr_time}')
                oc_row_number += rows_oc_df
                oc_row_number += 1
                oc.range(f'G{oc_row_number}' + ':' + f'T{oc_row_number}').color = COLOR_GREY               
                oc.range(f'G{oc_row_number}').value = df                            
                oc.range(f'F{oc_row_number}').value = oc_curr_time
                #oc.range(f'F{oc_row_number}').autofit()
                oc.range(f'F{oc_row_number}').font.bold = True
                oc.range(f'G{oc_row_number}' + ':' + f'T{oc_row_number}').font.bold = True

            if oc_row_number == 1 or (oc_row_number > 1 and duration is not None and duration.total_seconds() > 0):
                oc_prev_time = oc_curr_time
        else:
            logger.error(f'Error getting Options Data - Either Options DataFrame is Null or Expiry date is not entered')
            time.sleep(5)
            logger.debug("Trying to connect again...")
            nse = NSE()
            continue            
    ####################### OptionChain Ends ###########################

    ####################### EquityData Starts ###########################
    ind_sym, eq_sym = eq.range("E2").value, eq.range("E3").value
    if pre_ind_sym != ind_sym:
        eq_sym = None
        eq.range("I1:AD510").value = eq.range("D5:H30").value = None
        eq.range("E1").value = eq.range("G1").value = None
        eq.range("E3").value = eq.range("G2").value = None
        eq_row_number = 1
        eq_prev_time = eq_curr_time = None
        #prev_time_1 = datetime.now()

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
            eq.range("E1").value = eq_df.loc[ind_sym,'lastUpdateTime']
            eq.range("G2").value = eq_df.loc[ind_sym,'lastPrice']            
            eq.range("G1").value = eq_df.iloc[0]['lastUpdateTime']
            eq_curr_time = eq.range("G1").value            
            #eq.range("G1").autofit()

            if eq_row_number == 1 and eq_df_flag:
                eq.range(f'I{eq_row_number}').value = eq_df
                eq_df_flag = False

            if eq_prev_time != None and eq_prev_time != eq_curr_time:
                eq_row_number += rows_eq_df
                eq.range(f'I{eq_row_number}' + ':' + f'Z{eq_row_number}').color = COLOR_GREY               
                eq.range(f'I{eq_row_number}').value = eq_df                            
                eq.range(f'G{eq_row_number}').value = eq_curr_time
                eq.range(f'G{eq_row_number}').font.bold = True
                eq.range(f'I{eq_row_number}' + ':' + f'Z{eq_row_number}').font.bold = True

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
                    bid_list = ask_list = trd_data = []
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
                            break

                    bid_df = pd.DataFrame(bid_list)
                    bid_df.rename(columns={"price":"Bid Price","quantity":"Bid Quantity"},inplace=True)
                    ask_df = pd.DataFrame(ask_list)
                    ask_df.rename(columns={"price":"Ask Price","quantity":"Ask Quantity"},inplace=True)              
 
                    bid_ask_df = pd.concat([bid_df,ask_df], axis=1)

                    trd_df = pd.DataFrame(trd_data).transpose()
                    eq.range("D5").value = trd_df
                    eq.range("E5").value = None
                    eq.range("F6").value = "Lakhs"
                    eq.range("F7").value = "₹ Cr"
                    eq.range("F8").value = "₹ Cr"
                    eq.range("F9").value = "₹ Cr"
                    eq.range("D16").options(pd.DataFrame, index=False).value = bid_ask_df
                    eq.range("D22").value = "TotalBidQtyBuy"
                    eq.range("E22").value = tot_buy
                    eq.range("F22").value = "TotalBidQtySell"
                    eq.range("G22").value = tot_sell
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
    fd_sym = fd.range("E2").value
    if pre_fd_sym != fd_sym:
        fd.range("G1:AD100").value = None
        pre_fd_sym = fd_sym
    fd_df = None
    if fd_sym is not None:
        indices = True if fd_sym == "NIFTY" or fd_sym == "BANKNIFTY" else False
        fd_df = nse.futures_data(fd_sym, indices)
        if fd_df is not None:
            fd.range("G1").value = fd_df
        else:
            logger.error(f'Error getting Futures Data - Futures DataFrame is Null')
            time.sleep(5)
            logger.debug("Trying to connect again...")
            nse = NSE()
            continue
    ####################### FuturesData Ends ###########################




