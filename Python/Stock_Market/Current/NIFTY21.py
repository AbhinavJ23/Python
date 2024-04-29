import requests
import pandas as pd
import xlwings as xw
import time
import datetime as dt

sym = input("Enter the symbol (e.g NIFTY,BANKNIFTY,INFY,RELIANCE): ").upper()
file = xw.Book("OptionChain.xlsx")
sh1 = file.sheets("OptionChain")
sh2 = file.sheets("DerivedData")
sh1.clear()
sh2.clear()
derivatives_list = ['NIFTY','BANKNIFTY','FINNIFTY','NIFTYNXT50','MIDCPNIFTY']

def get_oc_data():
    if sym in derivatives_list:
        url = "https://www.nseindia.com/api/option-chain-indices?symbol="+sym
    else:
        url = "https://www.nseindia.com/api/option-chain-equities?symbol="+sym
    #print(url)
    headers = {"accept-encoding":"gzip, deflate, br, zstd",
           "accept-language":"en-US,en;q=0.9",
           "user-agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
}

    session = requests.Session()
    data = session.get(url, headers= headers).json()["records"]["data"]
    ocdata_ce = []
    ocdata_pe = []
    for i in data:
        for j,k in i.items():
            if j == "CE":
                info_ce = k
                ocdata_ce.append(info_ce)
            if j == "PE":
                info_pe = k
                ocdata_pe.append(info_pe)

    ce_df = pd.DataFrame(ocdata_ce)
    pe_df = pd.DataFrame(ocdata_pe)
    ce_df.columns += "_CE"
    pe_df.columns += "_PE"
    return pd.concat([ce_df, pe_df], axis = 1)

row_number=1
while True:
    try:
        df = pd.DataFrame(get_oc_data())
        underlying_value = df['underlyingValue_CE'].max()
        print(underlying_value)
        sh1.range("A1").value = df
        sum_oi_ce = df.groupby("expiryDate_CE")["openInterest_CE"].sum()
        sum_oi_pe = df.groupby("expiryDate_PE")["openInterest_PE"].sum()
        sum_lp_ce = df.groupby("expiryDate_CE")["lastPrice_CE"].sum()
        sum_lp_pe = df.groupby("expiryDate_PE")["lastPrice_PE"].sum()
        mul_ce = sum_oi_ce.multiply(sum_lp_ce)
        mul_pe = sum_oi_pe.multiply(sum_lp_pe)
        div_value = mul_ce.divide(mul_pe,fill_value=0)
        div_value.index.name = 'ExpiryDate'
        exp_list = []
        for i in div_value.index:
            exp_list.append(dt.datetime.strptime(i,'%d-%b-%Y').strftime('%Y-%m-%d'))
        df1 = pd.DataFrame(div_value.tolist(), index=exp_list, columns=['Value'])
        df1.index.name = 'ExpiryDate'
        pd.to_datetime(df1.index)
        sorted_idx = df1.index.sort_values()
        df1 = df1.loc[sorted_idx]
        sh2.range(f'A{row_number + 1}').api.Font.Bold = True
        #print(len(df1.columns))
        if row_number==1:
            sh2.range("A1").value = df1.transpose()
            sh2.range(f'A{row_number + 1}').value = dt.datetime.now()
            #print(sh2.range('A1').end('right').column)
            #sh2.range(f'A{num_col + 1}').value = "UnderlyingValue"
            #sh2.range(f'A{num_col + 1}').value = underlying_value             
            sh2.range("T1").value = "UnderlyingValue"
            sh2.range("T2").value = underlying_value
            sh2.range('1:1').api.Font.Bold = True
        else:
            sh2.range(f'A{row_number + 1}').value = dt.datetime.now()
            sh2.range(f'B{row_number + 1}:Z{row_number + 1}').number_format = "General"
            sh2.range(f'B{row_number + 1}').value = df1['Value'].to_list()
            sh2.range(f'T{row_number + 1}').value = underlying_value
            #sh2.range(f'B{num_col + 1}').value = underlying_value
        time.sleep(60)
        row_number += 1
    except:
        print("Retrying....")
        time.sleep(10)




