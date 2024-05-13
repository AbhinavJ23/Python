from NseIndia import NSE
import os,sys,time
import pandas as pd
import xlwings as xw
import dateutil.parser
import numpy as np
import pprint

nse = NSE()

## Creating new excel and adding sheets
if not os.path.exists("Nse_Data.xlsx"):
    try:
        wb = xw.Book()
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

oc.range('1:1').font.bold = True
oc.range('1:1').color = (211, 211, 211)
eq.range('1:1').font.bold = True
eq.range('1:1').color = (211, 211, 211)
fd.range('1:1').font.bold = True
fd.range('1:1').color = (211, 211, 211)

####################### Initializing OptionChain sheet #######################
oc.range("A:B").value = oc.range("D6:E19").value = oc.range("G1:V4000").value = None
df= pd.DataFrame({"FNO Symbol":["NIFTY", "BANKNIFTY"] + nse.equity_market_data("Securities in F&O", symbol_list=True)})
df = df.set_index("FNO Symbol", drop=True)
oc.range("A1").value = df
oc.range("D2").value, oc.range("D3").value = "Enter Symbol", "Enter Expiry"
pre_oc_sym = pre_oc_exp = ""
exp_list = []

######################### Initializing EquityData sheet #######################
eq.range("A:A").value = eq.range("D5:E30").value = eq.range("K1:AF4000").value = None
eq_df = pd.DataFrame({"Index Symbol":nse.equity_market_categories})
eq_df = eq_df.set_index("Index Symbol", drop=True)
eq.range("A1").value = eq_df
eq.range("D2").value, eq.range("D3").value = "Enter Index ", "Enter Equity"
pre_ind_sym = pre_eq_sym = ""

####################### Initializing FuturesData sheet #######################
fd.range("A:A").value = fd.range("G1:Z100").value = None
fd_df= pd.DataFrame({"FNO Symbol":["NIFTY", "BANKNIFTY"] + nse.equity_market_data("Securities in F&O", symbol_list=True)})
fd_df = fd_df.set_index("FNO Symbol", drop=True)
fd.range("A1").value = fd_df
fd.range("D2").value = "Enter Index/Equity"
pre_fd_sym = ""

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
            oc.range("G1").value = df
        except:
            pass

    ####################### OptionChain Ends ###########################

    ####################### EquityData Starts ###########################
    ind_sym, eq_sym = eq.range("E2").value, eq.range("E3").value
    if pre_ind_sym != ind_sym or pre_eq_sym != eq_sym:
        eq.range("K1:AF1000").value = eq.range("D6:H30").value = None
        if pre_ind_sym != ind_sym:
            eq.range("D6:H30").value = eq.range("E3").value = None
        pre_ind_sym = ind_sym
        pre_eq_sym = eq_sym 
    if ind_sym is not None:
        try:
            eq_df = nse.equity_market_data(ind_sym)
            eq_df.drop(["priority","date365dAgo","chart365dPath","date30dAgo","chart30dPath","chartTodayPath","series"],
                       axis=1,inplace=True)
            eq.range("K1").value = eq_df

            if eq_sym is not None:
                data = nse.equity_info(eq_sym, trade_info=True)
                #pprint.pprint(data)
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
                eq.range("D22").value = "TotalBuyQuantity"
                eq.range("E22").value = tot_buy
                eq.range("F22").value = "TotalSellQuantity"
                eq.range("G22").value = tot_sell

        except:
            pass
    ####################### EquityData Ends ###########################

    ####################### FuturesData Starts ###########################
    fd_sym = fd.range("E2").value
    if pre_fd_sym != fd_sym:
        fd.range("G1:Z100").value = None
        pre_fd_sym = fd_sym
    if fd_sym is not None:
        indices = True if fd_sym == "NIFTY" or fd_sym == "BANKNIFTY" else False
        try:
            fd_df = nse.futures_data(fd_sym, indices)
            #eq_df.drop(["priority","date365dAgo","chart365dPath","date30dAgo","chart30dPath","chartTodayPath","series"],
                       #axis=1,inplace=True)
            fd.range("G1").value = fd_df
        except:
            pass
    ####################### FuturesData Starts ###########################




