# -*- coding: utf-8 -*-
"""
Created on Tue Jan 11 11:49:42 2022

@author: mustafa.yarici
"""
import streamlit as st
import pandas as pd
from datetime import datetime
import numpy as np
import time
start_time = time.time()


import requests as req




#%%tarihler
date1="2021-12-01"
date2="2022-01-12"
pathresult='C:/sim/'+date1+'_'+date2+' result.xlsx'
#%%
ptf_url="https://seffaflik.epias.com.tr/transparency/service/market/mcp-smp"#PTF-#fba-#fbs-#Eşleşme miktarı -blok teklif miktarı
tklf_url='https://seffaflik.epias.com.tr/transparency/service/market/day-ahead-market-volume'
blk_url='https://seffaflik.epias.com.tr/transparency/service/market/amount-of-block'
ytp_url='https://seffaflik.epias.com.tr/transparency/service/consumption/load-estimation-plan'
rtm_url='https://seffaflik.epias.com.tr/transparency/service/production/wpp-generation-and-forecast'
#%%
tklf_df= pd.DataFrame(req.get(tklf_url,params={"startDate":date1,"endDate":date2}).json()['body']['dayAheadMarketVolumeList'])
tklf_df["date"]=pd.to_datetime(tklf_df["date"].str[0:-3], format='%Y-%m-%dT%H:%M:%S.%f')

ptf_df= pd.DataFrame(req.get(ptf_url,params={"startDate":date1,"endDate":date2}).json()['body']['mcpSmps'])
ptf_df["date"]=pd.to_datetime(ptf_df["date"].str[0:-3], format='%Y-%m-%dT%H:%M:%S.%f')

blk_df=pd.DataFrame(req.get(blk_url,params={"startDate":date1,"endDate":date2}).json()['body']['amountOfBlockList'])
blk_df["date"]=pd.to_datetime(blk_df["date"].str[0:-3], format='%Y-%m-%dT%H:%M:%S.%f')

ytp_df=pd.DataFrame(req.get(ytp_url,params={"startDate":date1,"endDate":date2}).json()['body']['loadEstimationPlanList'])
ytp_df["date"]=pd.to_datetime(ytp_df["date"].str[0:-3], format='%Y-%m-%dT%H:%M:%S.%f')

rtm_df=pd.DataFrame(req.get(rtm_url,params={"startDate":date1,"endDate":date2}).json()['body']['data'])
rtm_df["effectiveDate"]=pd.to_datetime(rtm_df["effectiveDate"].str[0:-3], format='%Y-%m-%dT%H:%M:%S.%f')
rtm_df.rename(columns ={'effectiveDate':'date'},inplace=True)

#%%
tklf_df['date'] = tklf_df['date'].dt.tz_localize(None)
ptf_df["date"]=ptf_df["date"].dt.tz_localize(None)
blk_df["date"]=blk_df["date"].dt.tz_localize(None)
ytp_df["date"]=ytp_df["date"].dt.tz_localize(None)
rtm_df["date"]=rtm_df["date"].dt.tz_localize(None)


#%%
result=tklf_df.merge(ptf_df,how='left',on='date').merge(blk_df,how='left',on='date').merge(ytp_df,how='left',on='date').merge(rtm_df,how='left',on='date')
result.drop(columns = ["mcpState","smp","smpDirection","blockBid","blockOffer","matchedBids","period",
                       "periodType","quantityOfAsk","quantityOfBid","matchedBids","volume","generation",
                       "quarter1","quarter2","quarter3","quarter4"],inplace=True)

result.rename(columns ={'matchedOffers':'Eşleşme','priceIndependentBid':'FBA','priceIndependentOffer':'FBS',
                        'mcp':'PTF','amountOfPurchasingTowardsBlock':'Blok Alış Teklif','amountOfPurchasingTowardsMatchBlock':'Blok Alış Eşleşme',
                        'amountOfSalesTowardsBlock':'Blok Satış Teklif','amountOfSalesTowardsMatchBlock':'Blok Satış Eşleşme','lep':'YTP','forecast':'Ritm'},inplace=True)
result.drop(columns = ['Blok Alış Teklif','Blok Satış Teklif'],inplace=True)

result['day']=result['date'].dt.day

result['hour']=result['date'].dt.hour

#%%burada csv at, site de yeniden oku düzenle
result.to_csv('C:/streamlitapp/heroku.csv',encoding='utf-8-sig',sep=";", decimal=",",index=None)  


#%%yeniden oku
result=pd.read_csv('C:/streamlitapp/heroku.csv',encoding='utf-8-sig',sep=";", decimal=",",index_col=False)
result['date']=pd.to_datetime(result['date'])

#%%

result['shortdate']=pd.to_datetime(result['date']).dt.strftime("%Y-%m-%d")


#%%
import datetime
min_date = datetime.date(result['date'][0].year,result['date'][0].month,result['date'][0].day)
max_date = datetime.date(result['date'].iloc[-1].year,result['date'].iloc[-1].month,result['date'].iloc[-1].day)
date1 = st.date_input('Gün 1',value=min_date,min_value=min_date,max_value=max_date)
date2 = st.date_input('Gün 2',value=min_date, min_value=min_date,max_value=max_date)

date1=str(date1)
date2=str(date2)

#%%
base1=result[result['shortdate']==date2].reset_index(drop=True)
base1.drop(columns = ['date','day','hour','shortdate'],inplace=True)

base2=result[result['shortdate']==date1].reset_index(drop=True)
base2.drop(columns = ['date','day','hour','shortdate'],inplace=True)

substrc=base1-base2


#%%
summary=substrc.copy()
summary = summary[["PTF","Eşleşme", "FBS","FBA","Blok Satış Eşleşme","Blok Alış Eşleşme","YTP","Ritm"]]
summary=summary.rename(columns=lambda x: x+' Değişim')
summary.insert(0,str(date2)+" PTF",base1['PTF'].reset_index(drop=True))

st.write("Gün2 - Gün1 Farkı")
st.dataframe(summary.style.format("{:.2f}"))

#%%


#%%
import xlsxwriter
from io import BytesIO,StringIO
def export_excel():
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine = 'xlsxwriter')
    summary.to_excel(writer, startrow = 1, merge_cells = False,header=False, index=False, sheet_name = date2)
    workbook = writer.book
    worksheet = writer.sheets[date2]

    (max_row, max_col) = summary.shape
    #column_settings = [{'header': column} for column in summary.columns]
    #worksheet.add_table('A1:J26', {'columns': column_settings,'autofilter': False})#
    #worksheet.set_column(1, max_col - 1, 16)
    #format = workbook.add_format()
    #format.set_bg_color('#eeeeee')
    #worksheet.set_column(0, 9, 28)
    writer.close()
    output.seek(0)
    exceldata=output.getvalue()
    return exceldata

df_xlsx = export_excel()
st.download_button(label='📥 Download Current Result',
                                data=df_xlsx ,
                                file_name= 'df_test.xlsx')


