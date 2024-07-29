from NseIndia import NSE
import os,sys,time
import pandas as pd
import xlwings as xw
import dateutil.parser
import numpy as np
import time
import logging
from datetime import datetime

####################### Initializing Logging Start #######################
logging.basicConfig(filename='Nse_Data_Historical_'+time.strftime('%Y%m%d%H%M%S')+'.log', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger()
####################### Initializing Logging End #######################

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
        wb.close()
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
oc.range('C1:C500').color = COLOR_GREY
oc.range('G1:G500').color = COLOR_GREY
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
    df= pd.DataFrame({"FNO Symbol":["NIFTY", "BANKNIFTY"] + nse.equity_market_data("Securities in F&O", symbol_list=True)})
except Exception as e:
    logger.critical(f'Error getting FNO symbols for Options Data - {e}')
df = df.set_index("FNO Symbol", drop=True)
oc.range("A1").value = df
oc.range("A1:A200").autofit()
oc.range("D2").value, oc.range("D3").value = "Enter Symbol ->", "Enter Expiry ->"
oc.range('D2').font.bold = True
oc.range('D3').font.bold = True
oc.range("D2:E3").autofit()
pre_oc_sym = pre_oc_exp = ""
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
eq.range("D2:E3").autofit()
eq.range('A35:G35').color = COLOR_GREY
pre_ind_sym = pre_eq_sym = ""
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
pre_fd_sym = ""
logger.debug("FuturesData sheet initialized")

####################### Initializing Global Variables #######################
row_number = 1
prev_time = curr_time = ""
equity_df_flag = True
prev_time_1 = datetime.now()

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
    if oc_sym is not None:
        indices = True if oc_sym == "NIFTY" or oc_sym == "BANKNIFTY" else False
        try:
            if not exp_list:
                for i in list(nse.options_data(oc_sym, indices)["expiryDate"]):
                    if dateutil.parser.parse(i).date() not in exp_list:
                        exp_list.append(dateutil.parser.parse(i).date())
                df = pd.DataFrame({"Expiry Date": [str(i) for i in sorted(exp_list)]})
                df = df.set_index("Expiry Date", drop=True)
                oc.range("B1").value = df
                oc.range("B1").autofit()
            df = nse.options_data(oc_sym, indices)
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

            oc.range("D6").value = [["Timestamp", timestamp],
                                    ["Spot LTP", underlying_value],
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
            oc.range("E6").autofit()
            oc.range("G1").value = df
        except Exception as e:
            logger.warning(f'Error getting Options Data - {e}')
            pass
    ####################### OptionChain Ends ###########################

    ####################### EquityData Starts ###########################
    ind_sym, eq_sym = eq.range("E2").value, eq.range("E3").value
    if pre_ind_sym != ind_sym:
        eq_sym = ""
        eq.range("I1:AD510").value = eq.range("D5:H30").value = None
        eq.range("E1").value = eq.range("G1").value = None
        eq.range("E3").value = eq.range("F2").value = eq.range("G2").value = None
        row_number = 1
        prev_time = curr_time = ""
        prev_time_1 = datetime.now()

    if pre_eq_sym != eq_sym:
        eq.range("D5:H30").value = None
        eq.range("F3").value = None
        eq.range("G3").value = None
    pre_ind_sym = ind_sym
    pre_eq_sym = eq_sym 
    if ind_sym is not None:
        try:
            logger.debug(f'value of row number is {row_number}')
            eq_df = nse.equity_market_data(ind_sym)
            eq_df.drop(["priority","date365dAgo","chart365dPath","date30dAgo","chart30dPath","chartTodayPath","series","identifier"],
                       axis=1,inplace=True)
            eq_df.index.name = 'symbol'            
            sorted_idx = eq_df.index.sort_values()
            eq_df = eq_df.loc[sorted_idx]
            rows_eq_df = len(eq_df.index)
            #eq.range("I1").value = eq_df
            eq.range("D1").value = "Current Index Time"
            eq.range("D1").autofit()
            eq.range("E1").value = eq_df.loc[ind_sym,'lastUpdateTime']
            eq.range("E1").autofit()
            eq.range("F1").value = "Current Equity Time"
            eq.range("F1").autofit()
            eq.range("F2").value = "Current Index Value"
            eq.range('F2').font.bold = True
            eq.range("G2").value = eq_df.loc[ind_sym,'lastPrice']
            #eq.range("G1").value = eq.range(f'V{row_number + 2}').value
            eq.range("G1").value = eq_df.iloc[0]['lastUpdateTime']
            curr_time = eq.range("G1").value            
            eq.range("G1").autofit()

            if row_number == 1 and equity_df_flag:
                eq.range(f'I{row_number}').value = eq_df
                equity_df_flag = False

            if prev_time != "" and prev_time != curr_time:
                row_number += rows_eq_df
                eq.range(f'I{row_number}' + ':' + f'Z{row_number}').color = COLOR_GREY               
                eq.range(f'I{row_number}').value = eq_df                            
                eq.range(f'G{row_number}').value = curr_time
                eq.range(f'G{row_number}').font.bold = True
                eq.range(f'I{row_number}' + ':' + f'Z{row_number}').font.bold = True

            prev_time = curr_time

            if eq_sym is not None:
                data = nse.equity_info(eq_sym, trade_info=True)
                bid_list = ask_list = trd_data = []
                tot_buy = tot_sell = 0
                eq.range("F3").value = "Current Equity Value"
                eq.range('F3').font.bold = True
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
            
        except Exception as e:
            logger.warning(f'Error getting Equity Data for time {curr_time} - {e}')
            pass
    ####################### EquityData Ends ###########################

    ####################### FuturesData Starts ###########################
    fd_sym = fd.range("E2").value
    if pre_fd_sym != fd_sym:
        fd.range("G1:AD100").value = None
        pre_fd_sym = fd_sym
    if fd_sym is not None:
        indices = True if fd_sym == "NIFTY" or fd_sym == "BANKNIFTY" else False
        try:
            fd_df = nse.futures_data(fd_sym, indices)
            fd.range("G1").value = fd_df
        except Exception as e:
            logger.warning(f'Error getting Futures Data - {e}')
            pass
    ####################### FuturesData Ends ###########################




