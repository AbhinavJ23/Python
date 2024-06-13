from NseIndia import NSE
import os,sys,time
import pandas as pd
import xlwings as xw
import dateutil.parser
import numpy as np
import time
import logging

####################### Initializing Logging Start #######################
logging.basicConfig(filename='Nse_Data_'+time.strftime('%Y%m%d')+'.log', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger()
####################### Initializing Logging End #######################

nse = NSE()

## Creating new excel and adding sheets
if not os.path.exists("Nse_Data.xlsx"):
    try:
        wb = xw.Book()
        wb.sheets.add("MaxVolumeTurnover")
        wb.sheets.add("SpotTurnover")
        wb.sheets.add("SpotVolume")
        wb.sheets.add("SpotPrice")
        wb.sheets.add("FuturesData")
        wb.sheets.add("EquityData")
        wb.sheets.add("OptionChain")
        wb.save("Nse_Data.xlsx")
        wb.close()
        logger.debug("Created Excel - Nse_Data.xlsx")
    except Exception as e:
        logger.critical(f'Error Creating Excel - {e}')
        sys.exit()

wb = xw.Book("Nse_Data.xlsx")
oc = wb.sheets("OptionChain")
eq = wb.sheets("EquityData")
fd = wb.sheets("FuturesData")
sp = wb.sheets("SpotPrice")
sv = wb.sheets("SpotVolume")
st = wb.sheets("SpotTurnover")
mv = wb.sheets("MaxVolumeTurnover")

####################### Initializing Constants #######################
COLOR_GREY = (211, 211, 211)
COLOR_GREEN = (0, 255, 0)
COLOR_RED = (255, 0, 0)
COLOR_YELLOW = (255, 255, 0)
MAX_VOLUME_PERCENT_DIFF = 500
MAX_TURNOVER_PERCENT_DIFF = 250
MAX_TURNOVER_VALUE_DIFF = 50000000
ONE_CRORE = 10000000

####################### Initializing Excel Sheets #######################
oc.range('1:1').font.bold = True
oc.range('1:1').color = COLOR_GREY
oc.range('C1:C500').color = COLOR_GREY
oc.range('G1:G500').color = COLOR_GREY
oc.range('C1').column_width = 2
oc.range('G1').column_width = 2
eq.range('1:1').font.bold = True
eq.range('1:1').color = COLOR_GREY
eq.range('B1:C500').color = COLOR_GREY
eq.range('B1').column_width = 1
eq.range('C1').column_width = 1
eq.range('H1').column_width = 2
eq.range('H1:H500').color = COLOR_GREY
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
oc.range("D2").value, oc.range("D3").value = "Enter Symbol", "Enter Expiry"
oc.range("D2:E3").autofit()
pre_oc_sym = pre_oc_exp = ""
exp_list = []
logger.debug("OptionChain sheet initialized")

######################### Initializing EquityData sheet #######################
eq.range("A:A").value = eq.range("D5:E30").value = eq.range("I1:AD100").value = None
eq_df = pd.DataFrame({"Index Symbol":nse.equity_market_categories})
eq_df = eq_df.set_index("Index Symbol", drop=True)
eq.range("A1").value = eq_df
eq.range("A1:A50").autofit()
eq.range("D2").value, eq.range("D3").value = "Enter Index ", "Enter Equity"
eq.range("D2:E3").autofit()
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
fd.range("D2").value = "Enter Index/Equity"
fd.range("D2").autofit()
pre_fd_sym = ""
logger.debug("FuturesData sheet initialized")

####################### Initializing Global Variables #######################
row_number = 1
col_number = 2
col_number_1 = 1
prev_temp_dict = {}
prev_time = curr_time = ""
prev_vol = curr_vol = []
prev_vol_diff = curr_vol_diff = []
prev_price = curr_price = []
prev_price_diff = curr_price_diff = []
prev_turn = curr_turn = []
prev_turn_diff = curr_turn_diff = []

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

############################# Start - Function to create Spot sheets #############################
def create_spot_sheets(df,sh_type,time,row_number,prev_spot,curr_spot,prev_spot_diff,curr_spot_diff,col_number_1,stock_list):
    logger.debug(f'Printing Spot Sheet - {sh_type} for {time} and row {row_number}')
    if sh_type == "Price":
        spot_df = df[["lastPrice"]]
        sh = sp
    elif sh_type == "Volume":
        spot_df = df[["totalTradedVolume"]]
        sh = sv
    elif sh_type == "Turnover":
        spot_df = df[["totalTradedValue"]]
        sh = st
    else:
        logger.error(f'Error! Unexpected Input - {sh_type}')

    sh.range('1:1').color = COLOR_GREY
    sh.range(f'A{row_number + 1}').font.bold = True
    
    spot_df1 = spot_df.transpose()
    col_number = 2
    iter = 0
    per_diff = 0    
    vol_list = []
    turn_list = []       
    for col_name in spot_df1:
        spot_value = spot_df1[col_name].values
        if spot_value is None:
            logger.debug(f'For Time {time} - In {sh_type} sheet for {col_name}, value is None, hence setting it to zero')
            spot_value = np.zeros(1)
        if row_number == 1:
            sh.range(f'{get_col_name(col_number)}' + str(row_number)).options(index=False).value = spot_df1[col_name]            
            sh.range(f'{get_col_name(col_number+1)}' + str(row_number)).value = "Difference"
            if sh_type in ("Volume","Turnover"):
                sh.range(f'{get_col_name(col_number+2)}' + str(row_number)).value = "% Change"
            sh.range('1:1').font.bold = True
            sh.range("A1:ZZ1").autofit()
            sh.range(f'A{row_number + 1}').value = time
            prev_spot.append(spot_value)
        else:                  
            sh.range(f'A{row_number + 1}').value = time.strftime('%H:%M:%S')
            sh.range(f'{get_col_name(col_number)}'+ str(row_number+1)).value = spot_value
            curr_spot.append(spot_value)            
            if curr_spot[iter] is not None and prev_spot[iter] is not None:
                val_diff = curr_spot[iter] - prev_spot[iter]                
            else:
                logger.debug(f'For Time {time} - In {sh_type} sheet for {col_name},  current value is {curr_spot[iter]} and previous value is {prev_spot[iter]}, hence setting its difference to zero')
                val_diff = np.zeros(1)
            if sh_type in ("Volume","Turnover") and val_diff < 0:
                ## This indicates some data issue from NSE
                logger.warning("Potential Data Issue from NSE!!")
                logger.debug(f'For Time {time} - In {sh_type} sheet for {col_name}, current value can not be less than previous value. Hence setting difference between current and previous value as zero')
                val_diff = np.zeros(1)
            sh.range(f'{get_col_name(col_number+1)}'+ str(row_number+1)).value = val_diff
            if row_number == 2:
                prev_spot_diff.append(val_diff)
                if sh_type == "Price":
                    if val_diff > 0:
                        sh.range(f'{get_col_name(col_number+1)}'+ str(row_number+1)).color = COLOR_GREEN
                    elif val_diff < 0:
                        sh.range(f'{get_col_name(col_number+1)}'+ str(row_number+1)).color = COLOR_RED
            else:
                curr_spot_diff.append(val_diff)
                if sh_type in ("Volume","Turnover"):
                    if prev_spot_diff[iter] != 0:
                        ## Calculate percentage difference only if denominator is not zero
                        per_diff = ((curr_spot_diff[iter] - prev_spot_diff[iter])*100)/prev_spot_diff[iter]
                        sh.range(f'{get_col_name(col_number+2)}'+ str(row_number+1)).value = per_diff
                    else:
                        logger.warning("Division by Zero occurred!!")
                        logger.debug(f'For Time {time} - Avoiding division by zero in {sh_type} sheet for {col_name}')
                        sh.range(f'{get_col_name(col_number+2)}'+ str(row_number+1)).value = "NA"
                        per_diff = np.zeros(1)
                        sh.range(f'{get_col_name(col_number+2)}'+ str(row_number+1)).color = COLOR_YELLOW

                if sh_type == "Volume":
                    if per_diff > MAX_VOLUME_PERCENT_DIFF:
                        sh.range(f'{get_col_name(col_number+2)}'+ str(row_number+1)).color = COLOR_GREEN
                    elif per_diff < -MAX_VOLUME_PERCENT_DIFF:
                        sh.range(f'{get_col_name(col_number+2)}'+ str(row_number+1)).color = COLOR_RED

                    if per_diff > MAX_VOLUME_PERCENT_DIFF or per_diff < -MAX_VOLUME_PERCENT_DIFF:                        
                        stock_list.append(col_name)                     
                        vol_list.append(per_diff)

                elif sh_type == "Turnover":
                    if per_diff > MAX_TURNOVER_PERCENT_DIFF:
                        sh.range(f'{get_col_name(col_number+2)}'+ str(row_number+1)).color = COLOR_GREEN
                    elif per_diff < -MAX_TURNOVER_PERCENT_DIFF:
                        sh.range(f'{get_col_name(col_number+2)}'+ str(row_number+1)).color = COLOR_RED
                    
                    temp_val_diff = val_diff/ONE_CRORE
                    if col_name in stock_list:                                             
                        turn_list.append(temp_val_diff)                       
                    elif col_name not in nse.equity_market_categories and val_diff > MAX_TURNOVER_VALUE_DIFF:
                        stock_list.append(col_name)
                        turn_list.append(temp_val_diff)

                elif sh_type == "Price":
                    if val_diff > 0:
                        sh.range(f'{get_col_name(col_number+1)}'+ str(row_number+1)).color = COLOR_GREEN
                    elif val_diff < 0:
                        sh.range(f'{get_col_name(col_number+1)}'+ str(row_number+1)).color = COLOR_RED               
        iter += 1
        col_number += 3

    ############### Start - Logic to print max volume and max turnover ############### 
    if sh_type in ("Volume","Turnover") and row_number >= 2:
        mv.range(f'{get_col_name(col_number_1)}' + '1').value = time.strftime("%H:%M:%S")
        mv.range(f'{get_col_name(col_number_1)}' + '1').font.bold = True 

        if sh_type == "Volume":
            if stock_list:        
                logger.debug(f'For MaxVolume, Stock List is - {stock_list}')
                logger.debug(f'For MaxVolume, Volume List is - {vol_list}')
                temp_vol_df = pd.DataFrame(vol_list, index=stock_list, columns=['Vol%Diff'])        
                mv.range(f'{get_col_name(col_number_1)}' + '2').value = temp_vol_df
            else:
                mv.range(f'{get_col_name(col_number_1+1)}' + '2').value = "Vol%Diff"
            mv.range(f'{get_col_name(col_number_1)}' + '2').value = "Name"
            mv.range(f'{get_col_name(col_number_1)}' + '2').font.bold = True
            mv.range(f'{get_col_name(col_number_1+1)}' + '2').font.bold = True            
            logger.debug("MaxVolume Printed")
        elif sh_type == "Turnover":
            if turn_list:
                stock_list.sort()        
                logger.debug(f'For MaxTurnover, Stock List is - {stock_list}')
                logger.debug(f'For MaxTurnover, Turnover List is - {turn_list}')
                temp_turn_df = pd.DataFrame(turn_list, index=stock_list, columns=['Turnover(₹ Cr)'])  
                mv.range(f'{get_col_name(col_number_1+2)}' + '2').value = temp_turn_df
            else:
                mv.range(f'{get_col_name(col_number_1+3)}' + '2').value = "Turnover(₹ Cr)"
            mv.range(f'{get_col_name(col_number_1+2)}' + '2').value = "Name"
            mv.range(f'{get_col_name(col_number_1+2)}' + '2').font.bold = True
            mv.range(f'{get_col_name(col_number_1+3)}' + '2').font.bold = True
            mv.range(f'{get_col_name(col_number_1+3)}' + '2').autofit()
            logger.debug("MaxTurnover Printed")
    ############### End - Logic to print max volume and max turnover ##################           
############################# End - Function to create Spot sheets #############################

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
            logger.critical(f'Error getting Options Data - {e}')
    ####################### OptionChain Ends ###########################

    ####################### EquityData Starts ###########################
    ind_sym, eq_sym = eq.range("E2").value, eq.range("E3").value
    if pre_ind_sym != ind_sym: #or pre_eq_sym != eq_sym:
        eq.range("I1:AD100").value = eq.range("D6:H30").value = eq.range("E3").value = None
        sv.clear()
        sp.clear()
        st.clear()
        mv.clear()
        row_number = 1
        col_number = 2
    if pre_eq_sym != eq_sym:
        eq.range("D6:H30").value = None
    pre_ind_sym = ind_sym
    pre_eq_sym = eq_sym 
    if ind_sym is not None:
        try:
            eq_df = nse.equity_market_data(ind_sym)
            eq_df.drop(["priority","date365dAgo","chart365dPath","date30dAgo","chart30dPath","chartTodayPath","series","identifier"],
                       axis=1,inplace=True)
            eq_df.index.name = 'symbol'            
            sorted_idx = eq_df.index.sort_values()
            eq_df = eq_df.loc[sorted_idx]
            eq.range("I1").value = eq_df
            eq.range("D1").value = "Index Timestamp"
            eq.range("D1").autofit()
            eq.range("E1").value = eq.range("V36").value
            eq.range("E1").autofit()
            eq.range("F1").value = "Equity Timestamp"
            eq.range("F1").autofit()
            eq.range("G1").value = eq.range("V3").value
            curr_time = eq.range("G1").value
            eq.range("G1").autofit()
            if eq_sym is not None:
                data = nse.equity_info(eq_sym, trade_info=True)
                bid_list = ask_list = trd_data = []
                tot_buy = tot_sell = 0

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

            ####################### Start - Spot Data (Price,Volume,Turnover) ###########################
            if prev_time != curr_time:
                stock_list = []
                vol_flag = False
                turn_flag = False
                create_spot_sheets(eq_df,"Price",curr_time,row_number,prev_price,curr_price,prev_price_diff,curr_price_diff,col_number_1,stock_list)                               
                create_spot_sheets(eq_df,"Volume",curr_time,row_number,prev_vol,curr_vol,prev_vol_diff,curr_vol_diff,col_number_1,stock_list)
                create_spot_sheets(eq_df,"Turnover",curr_time,row_number,prev_turn,curr_turn,prev_turn_diff,curr_turn_diff,col_number_1,stock_list)
                if row_number >= 2:
                    col_number_1 += 4
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
            ####################### End - Spot Data (Price,Volume,Turnover) ###########################               
        except Exception as e:
            logger.critical(f'Error getting Equity Data - {e}')
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
            logger.critical(f'Error getting Futures Data - {e}')
    ####################### FuturesData Ends ###########################




