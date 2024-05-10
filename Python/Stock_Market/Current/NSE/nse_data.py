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

## Initializing OptionChain sheet
oc.range("A:B").value = oc.range("D6:E19").value = oc.range("G1:V4000").value = None
df= pd.DataFrame({"FNO Symbol":["NIFTY", "BANKNIFTY"] + nse.equity_market_data("Securities in F&O", symbol_list=True)})
df = df.set_index("FNO Symbol", drop=True)
oc.range("A1").value = df
oc.range("D2").value, oc.range("D3").value = "Enter Symbol", "Enter Expiry"
pre_oc_sym = pre_oc_exp = ""
exp_list = []

## Initializing EquityData sheet
eq.range("A:A").value = None
eq_df = pd.DataFrame({"Index Symbol":nse.equity_market_categories})
eq_df = eq_df.set_index("Index Symbol", drop=True)
eq.range("A1").value = eq_df
eq.range("D2").value, oc.range("D3").value = "Enter Index ", "Enter Equity"
pre_ind_sym = pre_eq_sym = ""
#eq_list = []

print("Excel is starting up....")

while True:
    time.sleep(1)

    ## OptionChain
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
    
    ## EquityData
    ind_sym, eq_sym = eq.range("E2").value, eq.range("E3").value
    if pre_ind_sym != ind_sym or pre_eq_sym != eq_sym:
        eq.range("G1:AF1000").value = None
        if pre_ind_sym != ind_sym:
            eq.range("B:B").value = eq.range("D6:E19").value = None
            #eq_list = []
        pre_ind_sym = ind_sym
        pre_eq_sym = eq_sym 
    if ind_sym is not None:
        try:
            #if not eq_list:
            eq_df = nse.equity_market_data(ind_sym)
            #eq_df = eq_df[["symbol","identifier","open","dayHigh","dayLow","lastPrice","previousClose","change","pChange","ffmc",
            #               "yearHigh","yearLow","totalTradedVolume","totalTradedValue","lastUpdateTime","nearWKH","nearWKL",
            #               "perChange365d","perChange30d"]]
            #eq_df.set_index("symbol", drop=True, inplace=True)
            eq_df.drop(["priority","date365dAgo","chart365dPath","date30dAgo","chart30dPath","chartTodayPath","series"],
                       axis=1,inplace=True)
            eq.range("G1").value = eq_df
        except:
            pass
     




