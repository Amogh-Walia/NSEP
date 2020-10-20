import requests
import xlwings as xw
import pandas as pd
from time import sleep
from datetime import datetime
import os
import json
import winsound

def beep():
    print();print();print()
    print('error prevented')
    print();print();print()
    frequency = 2500  # Set Frequency To 2500 Hertz
    duration = 5000  # Set Duration To 1000 ms == 1 second
    winsound.Beep(frequency,duration)

print('-----------------------------------------------librairies loaded-----------------------------------------------')
pd.set_option('display.width', 1500)
pd.set_option('display.max_columns', 75)
pd.set_option('display.max_rows', 1500)



expiry_list =[]
file1 = open('Nifty expiry.txt','r')
for i in file1.readlines():
    if i[0] == '#' :        
        pass
    else:
        expiry_list.append(i[:-1])

expiry_list = expiry_list[:-1]
print(expiry_list)
excel_file = "nifty2.xlsm"

wb = xw.Book(excel_file)

symbol = ''
'''
    if page in sheet_list:
        pass
    else:
        wb.sheets.add(page)
'''
print('-----------------------------------------------constants loaded-----------------------------------------------')

def fetch_oi(df,expiry,mp_df,mp_list):
    symbol = 'NIFTY'
    sheet_list = []
    page = str(expiry)

    for i in wb.sheets:
        sheet_list.append(str(i)[19:-1])
#here
    max_retries = 3
    sheet = wb.sheets(page)
    tries = 1


    while tries <= max_retries:
        
        headers = {'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36",
                  "Accept-Language": "en-US,en;q=0.5","Accept-Encoding": "gzip, deflate"
                  }
        cookie_dict = {'bm_sv':'D4B12FF7ED731A7DA9A7FF5F2E82F219~Fkgs+aQbSMjZbEVeVewalaSResOVRIT/Qv060V57sp78r3VeIqsmcl/Cnbtzg9RVWnJLK2yFjFYY2ND5pPgXZm14O7QyReHh25W2KoL9iLxQPsZiT8C+J18mrsIwPOP9xnW0QlKS4GyDNQdWpPrNImIvavivzU/kB3QYa2XNLCI='
                                      
                       }
        session = requests.session()


        for cookie in cookie_dict:
            session.cookies.set(cookie, cookie_dict[cookie])

        url = "https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY"
        try:
            r = session.get(str(url), headers = headers).json()
            
        except:
            print('connection error')

            break
        try:
            
            if expiry:
                '''
                ce_values = [data['CE'] for data in r['records']['data'] if 'CE' in data and str(data['expiryDate']).lower( )== str(expiry).lower()]
                pe_values = [data['PE'] for data in r['records']['data'] if 'PE' in data and str(data['expiryDate']).lower( )== str(expiry).lower()]
                '''
                ce_values = []
                pe_values = []

                for o in r['records']['data']:
                    strike = 0
                    ide = ''

                    pe_written = False
                    ce_written = False
                    if 'CE' in o:
                        
                        data = o['CE']
                        #print(str(data['expiryDate']).lower() +'    '+ str(expiry).lower())
                        if str(data['expiryDate']).lower() == str(expiry).lower():
                            #print(1)
                            ce_written = True
                            strike =data['strikePrice']
                            ide = data['identifier']
                            ce_values.append(data)
                           # print(strike)
                    if 'PE' in o :

                        
                        data = o['PE']
                        if str(data['expiryDate']).lower() == str(expiry).lower():
                           #print('2')
                            pe_written = True
                            strike =data['strikePrice']
                            ide = data['identifier']
                            pe_values.append(data)
                            #print(strike)
                            
                    pe_value_strike = []
                    ce_value_strike = []
                    for ib in pe_values:
                        pe_value_strike.append(ib['strikePrice'])

                    for ib in ce_values:
                        ce_value_strike.append(ib['strikePrice'])

                    if not pe_written and strike not in pe_value_strike:
                        data1 = {'strikePrice': strike, 'expiryDate': expiry, 'underlying': symbol, 'identifier':ide, 'openInterest': 0, 'changeinOpenInterest': 0, 'pchangeinOpenInterest': 0, 'totalTradedVolume': 0, 'impliedVolatility': 0, 'lastPrice': 0, 'change': 0, 'pChange': 0, 'totalBuyQuantity': 10, 'totalSellQuantity': 0, 'bidQty': 0, 'bidprice': 0, 'askQty': 0, 'askPrice': 0, 'underlyingValue': 0}
                        pe_values.append(data1)

                    if not ce_written and strike not in ce_value_strike:
                        data1 = {'strikePrice': strike, 'expiryDate': expiry, 'underlying': symbol, 'identifier':ide, 'openInterest': 0, 'changeinOpenInterest': 0, 'pchangeinOpenInterest': 0, 'totalTradedVolume': 0, 'impliedVolatility': 0, 'lastPrice': 0, 'change': 0, 'pChange': 0, 'totalBuyQuantity': 10, 'totalSellQuantity': 0, 'bidQty': 0, 'bidprice': 0, 'askQty': 0, 'askPrice': 0, 'underlyingValue': 0}
                        ce_values.append(data1)





            else:
                
                '''            
                ce_values = [data['CE'] for data in r['filtered']['data'] if 'CE' in data]
                pe_values = [data['PE'] for data in r['filtered']['data'] if 'PE' in data]
                '''
                ce_values = []
                pe_values = []

                for o in r['filtered']['data']:

                    data = o['CE']
                    if 'CE' in data:
                        ce_values.append(data['CE'])
                

                for o in r['filtered']['data']:
                    data = o['PE']
                    if 'PE' in data :
                        pe_values.append(data['PE'])
        except:
            print('error averted')
            beep()
            break
        for ce in ce_values:
            
            strike = ce['strikePrice']
            if strike == '':
                pass
            for pe in pe_values:
                if pe['strikePrice'] == strike:
                    if ce['openInterest'] == 0:
                        pcr = 'Null'
                    else:
                        pcr = float(pe['openInterest'])/float(ce['openInterest'])
                    if ce['changeinOpenInterest'] == 0:
                        changepcr = 'Null'
                    else:
                        changepcr = float(pe['changeinOpenInterest'])/float(ce['changeinOpenInterest'])

                    call_analysis = ''
                    pcall_analysis = ''
                    if ce['change'] >= 0 and ce['changeinOpenInterest'] >=0:
                        call_analysis =  'Fresh Long'
                    if ce['change'] >=0 and ce['changeinOpenInterest'] <=0:
                        call_analysis =  'Short Covering'
                    if ce['change'] <= 0 and ce['changeinOpenInterest'] >=0:
                        call_analysis =  'Fresh Short'
                    if ce['change'] <= 0 and ce['changeinOpenInterest'] <=0:
                        call_analysis =  'Long Unwinding'

                        
                    if pe['change'] >= 0 and pe['changeinOpenInterest'] >=0:
                        pcall_analysis =  'Fresh Long'
                    if pe['change'] >= 0 and pe['changeinOpenInterest'] <=0:
                        pcall_analysis =  'Short Covering'
                    if pe['change'] <= 0 and pe['changeinOpenInterest'] >=0:
                        pcall_analysis =  'Fresh Short'
                    if pe['change'] <= 0 and pe['changeinOpenInterest'] <=0:
                        pcall_analysis =  'Long Unwinding'


                    if ce['change'] == 0 and ce['changeinOpenInterest'] ==0:
                        call_analysis =  'NULL'
                    if pe['change'] == 0 and pe['changeinOpenInterest'] ==0:
                        pcall_analysis =  'NULL'


                    

                    pe['PCR'] = pcr
                    pe['Change in PCR'] = changepcr
                    pe['Sum of oi'] = pe['openInterest']+ce['openInterest']
                    pe['Call Analysis'] = call_analysis
                    pe['Put Analysis'] = pcall_analysis
                    pe['Call Premium turnover']= ce['lastPrice']*ce['changeinOpenInterest']
                    pe['Put Premium turnover']= pe['lastPrice']*pe['changeinOpenInterest']
                    
                   # print(ce['underlyingValue'])
                    #print(ce['strikePrice'])
                    pe['Strike Price - underlying'] = abs(float(ce['underlyingValue'])-float(ce['strikePrice']))

                    pe['Differ IV'] = abs(float(ce['impliedVolatility'])-float(pe['impliedVolatility']))

        ce_data = pd.DataFrame(ce_values)
        pe_data = pd.DataFrame(pe_values)
        ce_data = ce_data.sort_values(['strikePrice'])
        pe_data = pe_data.sort_values(['strikePrice'])
        #sheet.clear_contents()

        sheet.range("A3").options(index=False,header = True).value = ce_data.drop(
            ['askPrice','askQty','bidQty','bidprice','expiryDate','identifier','totalBuyQuantity','totalSellQuantity','underlying'], axis = 1)[
             ['openInterest','changeinOpenInterest','pchangeinOpenInterest','impliedVolatility','lastPrice','change','pChange','totalTradedVolume','underlyingValue','strikePrice']]

        sheet.range("K3").options(index=False,header = True).value = pe_data.drop(
            ['askPrice','askQty','bidQty','bidprice','expiryDate','identifier','totalBuyQuantity','totalSellQuantity'], axis = 1)[
             ['strikePrice','openInterest','changeinOpenInterest','pchangeinOpenInterest','impliedVolatility','totalTradedVolume','lastPrice','change','pChange','underlying','Sum of oi','PCR','Change in PCR','Call Analysis','Put Analysis','underlyingValue','Strike Price - underlying','Differ IV','Put Premium turnover','Call Premium turnover']]



        Time = datetime.now().strftime('%H:%M')
        sheet.range("A1").value = str(Time)
        Time1 = '[ '+Time+' ]'
        print(Time1+'>>Data pasted Now Recording')
        break
        ce_data['type'] = 'CE'
        pe_data['type'] = 'PE'
        df1 =  pd.concat([ce_data,pe_data])
        
        if len(df_list) > 0:
            df1['Time'] = df_list[-1][0]['Time']

        if len(df_list) > 0 and df1.to_dict('records') == df_list[-1]:
            print(Time1+'>>Duplicate data Not recording')
            break

            
 
        df1['Time'] = Time
                     
        if not df.empty:
            df = df[
                [ "Time",
            "askPrice",
            "askQty",
            "bidQty",
            "bidprice",
            "change",
            "changeinOpenInterest",
            "expiryDate",
            "identifier",
            "impliedVolatility",
            "lastPrice",
            "openInterest",
            "pChange",
            "pchangeinOpenInterest",
            "strikePrice",
            "totalBuyQuantity",
            "totalSellQuantity",
            "totalTradedVolume",
            "type",
            "underlying",
            "underlyingValue"]]

            df1 = df1[
                [ "Time",
            "askPrice",
            "askQty",
            "bidQty",
            "bidprice",
            "change",
            "changeinOpenInterest",
            "expiryDate",
            "identifier",
            "impliedVolatility",
            "lastPrice",
            "openInterest",
            "pChange",
            "pchangeinOpenInterest",
            "strikePrice",
            "totalBuyQuantity",
            "totalSellQuantity",
            "totalTradedVolume",
            "type",
            "underlying",
            "underlyingValue"]]

            
        df = pd.concat([df,df1])
        
        df_list.append(df1.to_dict('records'))
        with open(oi_filename,'w') as files:
            files.write(json.dumps(df_list,indent = 4, sort_keys = True))




def main():
    print('-----------------------------------------------running secondary loop-----------------------------------------------')
    global df_list
    global oi_filename
    global mp_list
    global mp_df
    global mp_filename
    mp_list = []
    mp_df = pd.DataFrame()
    print('-----------------------------------------------global variables loaded-----------------------------------------------')
    for ia in expiry_list:

        print('Aquiring data for date:'+str(ia))
        df_list = []
        oi_filename = os.path.join("Files","oi_data_records_of_expiry_{0}_on_{1}.json".format(ia,datetime.now().strftime('%d%m%y')))
        mp_filename = os.path.join("Files","mp_data_records_of_expiry_{0}_on_{1}.json".format(ia,datetime.now().strftime('%d%m%y')))
        

        try:
            df_list = json.loads(open(oi_filename).read())
        except Exception as error:
                print("error reading data. Error : {0}".format(error))
                df_list = []
                

        if df_list:
                df =pd.DataFrame()
                for item in df_list:
                    df = pd.concat([df, pd.DataFrame(item)])
        else:
                df = pd.DataFrame()
        
        try:
            mp_list = json.loads(open(mp_filename).read())
            mp_df   = pd.DataFrame().from_dict(mp_list)

        except Exception as error:
            print("error reading data. Error : {0}".format(error))
            mp_list = []
            mp_df = pd.DataFrame()


        print('-----------------------------------------------Fetching data-----------------------------------------------')
        fetch_oi(df,ia,mp_df,mp_list)
def OptionCombined():
    url = "https://www.nseindia.com/api/liveEquity-derivatives?index=nse50_opt"

    excel_file = "nifty2.xlsm"

    wb = xw.Book(excel_file)


    print('-----------------------------------------------constants loaded-----------------------------------------------')
    headers = {'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36",
                      "Accept-Language": "en-US,en;q=0.5","Accept-Encoding": "gzip, deflate"
                      }
    cookie_dict = {'bm_sv':

                   'D4B12FF7ED731A7DA9A7FF5F2E82F219~Fkgs+aQbSMjZbEVeVewalaSResOVRIT/Qv060V57sp78r3VeIqsmcl/Cnbtzg9RVWnJLK2yFjFYY2ND5pPgXZm14O7QyReHh25W2KoL9iLxQPsZiT8C+J18mrsIwPOP9xnW0QlKS4GyDNQdWpPrNImIvavivzU/kB3QYa2XNLCI='              
                   }





    session = requests.session()

'''
    for cookie in cookie_dict:
        session.cookies.set(cookie, cookie_dict[cookie])
        
    try:
        r = session.get(str(url), headers = headers).json()
        run_sucess = True
    except:
        run_sucess = False
        print('connection error')
    if run_sucess == True:
        data = []
        for i in r['data']:
            data.append(i)
        data1 = pd.DataFrame(data)
    sheet = wb.sheets('option combined')
    sheet.range("A3").options(index=False,header = True).value = data1
'''
def future():
    url = "https://www.nseindia.com/api/liveEquity-derivatives?index=nse50_fut"

    excel_file = "nifty2.xlsm"

    wb = xw.Book(excel_file)


    print('-----------------------------------------------constants loaded-----------------------------------------------')
    headers = {'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36",
                      "Accept-Language": "en-US,en;q=0.5","Accept-Encoding": "gzip, deflate"
                      }
    cookie_dict = {'bm_sv':

                   '333BF02F4AC4272A7CB762987035BDE1~6V3MHeSeqBEK2b3ri7n9CrFJkAhbu16/CLO7wyMByVoUaCtrgWj9h391Z89eTqZhI7pnme6Bpu1SQa13xK0llhSEIsVpIciB1KldpAFY036sGqRSvXc5Fv+VQ8h9Q229ASUBKwYdAKogR5Z7EOai1hdYNteDTHVt53He+hdOX68='
                                  
                   }





    session = requests.session()


    for cookie in cookie_dict:
        session.cookies.set(cookie, cookie_dict[cookie])
        

    r = session.get(str(url), headers = headers).json()
    print('-----------------------------------------------------------------------------')
    lol = []
    for i in r['data']:
        cv = {}
        for j  in i:
            if j == 'meta':
                pass
                
            else:
               cv[j] = i[j]
            
        lol.append(cv)

    l = pd.DataFrame(lol)
    sheet = wb.sheets('future')
    sheet.range("A3").options(index=False,header = True).value = l


while True:
    headers = {'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36",
                  "Accept-Language": "en-US,en;q=0.5","Accept-Encoding": "gzip, deflate"
                  }
    cookie_dict = {'bm_sv': '91B0B0C57704155316E9CFD21E65333~/CzCOpuWPhk3L+Xq7JRzjS0+MjuvU2AvTVrpprjqoHv0YfY46s1UCHqPkEt0WFTfaqECtHRuNiiBgvrNV19K7/+K/H6faeO43gOtLqD0yRJ3GIgyQlPZOa4BRsezY+jI9OuHOJKAag7DxFCnQFsuRFp9HjgaUlNYmu6pNJpr31E='

                                      
                       }
    session = requests.session()


    for cookie in cookie_dict:
        session.cookies.set(cookie, cookie_dict[cookie])
    df_list = []
    mp_list = []
    start_time = datetime.now().strftime('%H:%M')
    sheet_list = []

    for it in wb.sheets:
        sheet_list.append(str(it)[19:-1])

    url = "https://www.nseindia.com/api/liveEquity-derivatives?index=stock_opt"
    try:
        session = requests.session()


        for cookie in cookie_dict:
            session.cookies.set(cookie, cookie_dict[cookie])
        r = session.get(str(url), headers = headers).json()

        run_sucess = True
        data = []
        for i in r['data']:
            i['Time'] = datetime.now().strftime('%H:%M')
            data.append(i)
        data1 = pd.DataFrame(data)
        sheet = wb.sheets('Most Active')

        sheet.range("A10").options(index=False,header = True).value = data1[['underlying','identifier','instrumentType','instrument','contract','expiryDate','optionType','strikePrice','lastPrice','change','pChange','volume','totalTurnover','value','premiumTurnOver','underlyingValue','openInterest','noOfTrades','Time']]
    except:
        print('Failed to get most active sheet')


        
    print('-----------------------------------------------------------------------------------------------------------------------')
    print('Initiating loop at: '+start_time)
    print('-----------------------------------------------------------------------------------------------------------------------')
    
    main()
    print('running option____________________________________________________________ ')
    OptionCombined()
    print('future option____________________________________________________________ ')
#    future()
    end_time = datetime.now().strftime('%H:%M')
    minutes = str(int(end_time[3:])-int(start_time[3:]))
    hours = str(int(end_time[:-3])-int(start_time[:-3]))
    print('-----------------------------------------------------------------------------------------------------------------------')
    print('Ending loop at: '+end_time)
    print('>>>Time Taken :'+hours+' hours and '+minutes+' minutes')
     
    print('-----------------------------------------------------------------------------------------------------------------------')
    sleep(30)
