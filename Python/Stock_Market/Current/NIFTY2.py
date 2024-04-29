import requests
import pandas as pd
import xlwings as xw
import time
import datetime as dt

sym = "NIFTY"
#sym = input("Enter the symbol (e.g NIFTY, INFY): ").upper()
file = xw.Book("OC_Data21.xlsx")
sh1 = file.sheets("Nifty")
sh2 = file.sheets("Data")
sh1.clear()
sh2.clear()

def get_oc_data():
    url = "https://www.nseindia.com/api/option-chain-indices?symbol="+sym
    #url = "https://www.nseindia.com/api/option-chain-equities?symbol="+sym
    headers = {"accept-encoding":"gzip, deflate, br, zstd",
           "accept-language":"en-US,en;q=0.9",
           "user-agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
}

    session = requests.Session()
    data = session.get(url, headers= headers).json()["records"]["data"]
    ocdata_ce = []
    ocdata_pe = []
    #print(data)
    for i in data:
        for j,k in i.items():
            #if j=="CE" or j=='PE':
            if j == "CE":
                info_ce = k
                ocdata_ce.append(info_ce)
            if j == "PE":
                info_pe = k
                ocdata_pe.append(info_pe)
                #info["instrumentType"]=j

    ce_df = pd.DataFrame(ocdata_ce)
    pe_df = pd.DataFrame(ocdata_pe)
    ce_df.columns += "_CE"
    pe_df.columns += "_PE"
    return pd.concat([ce_df, pe_df], axis = 1)

#print(ocdata)
row_number=1
while True:
    try:
        df = pd.DataFrame(get_oc_data())
        #print(df)
        underlying_value = df['underlyingValue_CE'].max()
        print(underlying_value)
        sh1.range("A1").value = df
        #for column in df:
            #if df['instrumentType'] == 'CE':
        sum_oi_ce = df.groupby("expiryDate_CE")["openInterest_CE"].sum()
        #print(sum_oi_ce)
        sum_oi_pe = df.groupby("expiryDate_PE")["openInterest_PE"].sum()
        #print(sum_oi_pe)
        sum_lp_ce = df.groupby("expiryDate_CE")["lastPrice_CE"].sum()
        #print(sum_lp_ce)
        sum_lp_pe = df.groupby("expiryDate_PE")["lastPrice_PE"].sum()
        #print(sum_lp_pe)
        #for i,j in sum_oi_ce.items():
            #print('expiryDate_CE: ', i, 'openInterest_CE: ', j)
        mul_ce = sum_oi_ce.multiply(sum_lp_ce)
        #print(mul_ce)
        mul_pe = sum_oi_pe.multiply(sum_lp_pe)
        div_value = mul_ce.divide(mul_pe,fill_value=0)
        div_value.index.name = 'ExpiryDate'
        #print(div_value)
        exp_list = []
        for i in div_value.index:
            exp_list.append(dt.datetime.strptime(i,'%d-%b-%Y').strftime('%Y-%m-%d'))
        #print(exp_list)
        #pd.to_datetime(div_value.index)
        #df2['ExpiryDate'] = pd.to_datetime(df2['ExpiryDate'])
        #sorted_idx = div_value.index.sort_values()
        #div_value = div_value.loc[sorted_idx]
        #div_value.index = pd.to_datetime(div_value.index)
        #div_value.sort_index()
        df1 = pd.DataFrame(div_value.tolist(), index=exp_list, columns=['Value'])
        #print(div_value.tolist())
        df1.index.name = 'ExpiryDate'
        #exp_list = []
        #for i in df1.index:
            #exp_list.append(dt.datetime.strptime(i,'%d-%b-%Y').strftime('%Y-%m-%d'))
        print(df1)
        #df2 = pd.DataFrame(df1['Value'].tolist(),index=exp_list,columns=['Value'])
        #df2.index.name = 'ExpiryDate'
        #print(df2)
        #df1.reset_index(level=0, inplace=True)
        pd.to_datetime(df1.index)
        #df1['ExpiryDate'] = pd.to_datetime(df1['ExpiryDate'])
        sorted_idx = df1.index.sort_values()
        df1 = df1.loc[sorted_idx]
        #df1.reset_index(level=0, inplace=True)
        print(df1)
        #print(type(df1['Value']))
        #pd.to_datetime(df1.index)
        #df1.sort_index()
        #dict_value = div_value.to_dict()
        #df1 = pd.DataFrame(div_value)
        #print(df1['Value'])
        #print(dict_value)
        if row_number==1:
            sh2.range("A1").value = df1.transpose()
            sh2.range(f'A{row_number + 1}').value = dt.datetime.now()
            sh2.range("V1").value = "UnderlyingValue"
            sh2.range("V2").value = underlying_value
            sh2.range('1:1').api.Font.Bold = True
            #sh2.range("A1").value = df1.index()
            #sh2.range("A1").value = dict_value
            #sh2.range("A2").value = dict_value
        else:
            sh2.range(f'A{row_number + 1}').value = dt.datetime.now()
            sh2.range(f'B{row_number + 1}:Z{row_number + 1}').number_format = "General"
            sh2.range(f'B{row_number + 1}').value = df1['Value'].to_list()
            sh2.range(f'V{row_number + 1}').value = underlying_value
            #sh2.range(f'A{row_number + 1}').value = list(float(dict_value.values()))
            #sh2.range(n+2,1).value = df1['Value'].transpose()
        time.sleep(60)
        row_number += 1
    except:
        print("Retrying....")
        time.sleep(10)




