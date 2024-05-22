from NseIndia import NSE
import os,sys,time
import pandas as pd
import xlwings as xw
import dateutil.parser
import numpy as np

nse = NSE()

## Creating new excel and adding sheets
if not os.path.exists("Nse_Data.xlsx"):
    try:
        wb = xw.Book()
        wb.sheets.add("SpotTurnover")
        wb.sheets.add("SpotVolume")
        wb.sheets.add("SpotPrice")
        wb.sheets.add("FuturesData")
        wb.sheets.add("EquityData")
        wb.sheets.add("OptionChain")
        wb.save("Nse_Data.xlsx")
        wb.close()
    except Exception as e:
        print(f'Error Creating Excel {e}')
        sys.exit()

wb = xw.Book("Nse_Data.xlsx")
oc = wb.sheets("OptionChain")
eq = wb.sheets("EquityData")
fd = wb.sheets("FuturesData")
sp = wb.sheets("SpotPrice")
sv = wb.sheets("SpotVolume")
st = wb.sheets("SpotTurnover")

oc.range('1:1').font.bold = True
oc.range('1:1').color = (211, 211, 211)
oc.range('C1:C500').color = (211, 211, 211)
oc.range('G1:G500').color = (211, 211, 211)
oc.range('C1').column_width = 2
oc.range('G1').column_width = 2
eq.range('1:1').font.bold = True
eq.range('1:1').color = (211, 211, 211)
eq.range('B1:C500').color = (211, 211, 211)
eq.range('B1').column_width = 1
eq.range('C1').column_width = 1
eq.range('H1').column_width = 2
eq.range('H1:H500').color = (211, 211, 211)
fd.range('1:1').font.bold = True
fd.range('1:1').color = (211, 211, 211)

####################### Initializing OptionChain sheet #######################
oc.range("A:B").value = oc.range("D6:E19").value = oc.range("G1:V4000").value = None
df= pd.DataFrame({"FNO Symbol":["NIFTY", "BANKNIFTY"] + nse.equity_market_data("Securities in F&O", symbol_list=True)})
df = df.set_index("FNO Symbol", drop=True)
oc.range("A1").value = df
oc.range("A1:A200").autofit()
oc.range("D2").value, oc.range("D3").value = "Enter Symbol", "Enter Expiry"
oc.range("D2:E3").autofit()
pre_oc_sym = pre_oc_exp = ""
exp_list = []

######################### Initializing EquityData sheet #######################
eq.range("A:A").value = eq.range("D5:E30").value = eq.range("I1:AD100").value = None
eq_df = pd.DataFrame({"Index Symbol":nse.equity_market_categories})
eq_df = eq_df.set_index("Index Symbol", drop=True)
eq.range("A1").value = eq_df
eq.range("A1:A50").autofit()
eq.range("D2").value, eq.range("D3").value = "Enter Index ", "Enter Equity"
eq.range("D2:E3").autofit()
pre_ind_sym = pre_eq_sym = ""

####################### Initializing FuturesData sheet #######################
fd.range("A:A").value = fd.range("G1:AD100").value = None
fd_df= pd.DataFrame({"FNO Symbol":["NIFTY", "BANKNIFTY"] + nse.equity_market_data("Securities in F&O", symbol_list=True)})
fd_df = fd_df.set_index("FNO Symbol", drop=True)
fd.range("A1").value = fd_df
fd.range("A1:A200").autofit()
fd.range("D2").value = "Enter Index/Equity"
fd.range("D2").autofit()
pre_fd_sym = ""

row_number = 1
col_number = 2
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
############################# End - Function to get excel column(A1,B1 etc) given a positive number #############################

############################# Start - Function to create Spot sheets #############################
def create_spot_sheets(df,sh_type,time,row_number,prev_spot,curr_spot,prev_spot_diff,curr_spot_diff):
    print("Starting Spot Sheet - ", sh_type," at ", time," for row ", row_number )
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
        print("Error! Unexpected Input - ", sh_type)

    sh.range('1:1').color = (211, 211, 211)
    sh.range(f'A{row_number + 1}').font.bold = True
    
    spot_df1 = spot_df.transpose()
    col_number = 2
    iter = 0
    per_diff = 0                
    for col_name in spot_df1:            
        if row_number == 1:
            sh.range(f'{get_col_name(col_number)}' + str(row_number)).options(index=False).value = spot_df1[col_name]            
            sh.range(f'{get_col_name(col_number+1)}' + str(row_number)).value = "Difference"
            if sh_type in ("Volume","Turnover"):
                sh.range(f'{get_col_name(col_number+2)}' + str(row_number)).value = "% Change"
            sh.range('1:1').font.bold = True
            sh.range("A1:ZZ1").autofit()
            sh.range(f'A{row_number + 1}').value = time
            prev_spot.append(spot_df1[col_name].values)
        else:                  
            sh.range(f'A{row_number + 1}').value = time
            sh.range(f'{get_col_name(col_number)}'+ str(row_number+1)).value = spot_df1[col_name].values
            curr_spot.append(spot_df1[col_name].values)
            val_diff = curr_spot[iter] - prev_spot[iter]
            sh.range(f'{get_col_name(col_number+1)}'+ str(row_number+1)).value = val_diff
            if row_number == 2:
                prev_spot_diff.append(val_diff)
                if sh_type in ("Price"):
                    if val_diff > 0:
                        sh.range(f'{get_col_name(col_number+1)}'+ str(row_number+1)).color = (0, 255, 0)
                    elif val_diff < 0:
                        sh.range(f'{get_col_name(col_number+1)}'+ str(row_number+1)).color = (255, 0, 0)   
            else:
                curr_spot_diff.append(val_diff)
                if sh_type in ("Volume","Turnover"):
                    if prev_spot_diff[iter] != 0:
                        per_diff = ((curr_spot_diff[iter] - prev_spot_diff[iter])*100)/prev_spot_diff[iter]
                        sh.range(f'{get_col_name(col_number+2)}'+ str(row_number+1)).value = per_diff
                    else:
                        sh.range(f'{get_col_name(col_number+2)}'+ str(row_number+1)).value = "NA"
                if sh_type in ("Volume"):
                    if per_diff > 500:
                        sh.range(f'{get_col_name(col_number+2)}'+ str(row_number+1)).color = (0, 255, 0)
                    elif per_diff < -500:
                        sh.range(f'{get_col_name(col_number+2)}'+ str(row_number+1)).color = (255, 0, 0)
                elif sh_type in ("Turnover"):
                    if per_diff > 250:
                        sh.range(f'{get_col_name(col_number+2)}'+ str(row_number+1)).color = (0, 255, 0)
                    elif per_diff < -250:
                        sh.range(f'{get_col_name(col_number+2)}'+ str(row_number+1)).color = (255, 0, 0)                   
                elif sh_type in ("Price"):
                    if val_diff > 0:
                        sh.range(f'{get_col_name(col_number+1)}'+ str(row_number+1)).color = (0, 255, 0)
                    elif val_diff < 0:
                        sh.range(f'{get_col_name(col_number+1)}'+ str(row_number+1)).color = (255, 0, 0)                  
        iter += 1
        col_number += 3                
############################# End - Function to create Spot sheets #############################

print("Excel is starting....")

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
        except:
            pass

    ####################### OptionChain Ends ###########################

    ####################### EquityData Starts ###########################
    ind_sym, eq_sym = eq.range("E2").value, eq.range("E3").value
    if pre_ind_sym != ind_sym: #or pre_eq_sym != eq_sym:
        eq.range("I1:AD100").value = eq.range("D6:H30").value = eq.range("E3").value = None
        sv.clear()
        sp.clear()
        st.clear()
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
            eq.range("E1").value = eq.range("V2").value
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
                eq.range("D22").value = "TotalBuyQty"
                eq.range("E22").value = tot_buy
                eq.range("F22").value = "TotalSellQty"
                eq.range("G22").value = tot_sell

            ####################### Start Spot Data (Vol,Price,Turnover) ###########################            
            if prev_time != curr_time:
                create_spot_sheets(eq_df,"Price",curr_time,row_number,prev_price,curr_price,prev_price_diff,curr_price_diff)
                create_spot_sheets(eq_df,"Volume",curr_time,row_number,prev_vol,curr_vol,prev_vol_diff,curr_vol_diff)
                create_spot_sheets(eq_df,"Turnover",curr_time,row_number,prev_turn,curr_turn,prev_turn_diff,curr_turn_diff)
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
            ####################### End - Spot Data (Vol,Price,Turnover) ###########################               
        except:
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
        except:
            pass
    ####################### FuturesData Ends ###########################




