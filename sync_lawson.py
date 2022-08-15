# -*- coding: utf-8 -*-
"""
Created on Mon Jun 14 10:16:49 2021

@author: dule4
"""
import pandas as pd
import pyodbc

lawson  = pyodbc.connect('Driver={Oracle in OraClient12Home1};DBQ=LAWPROD;Uid=dule4;Pwd=Phubai33')
lw_query=("SELECT COMPANY,LOCATION,DOC_TYPE,SYSTEM_CD,DOCUMENT,LINE_NBR,UPDATE_DATE,ITEM,TRAN_UOM,SOH_QTY,QUANTITY,UNIT_COST,LAST_UPDATE_BY FROM LAWPROD.ICTRANS i WHERE i.COMPANY ='3848' AND HIST_YEAR='2022' AND HIST_PERIOD>='6'")
print(lw_query)
# lw_query=("SELECT * FROM LAWPROD.ICTRANS i WHERE i.COMPANY ='3848' AND HIST_YEAR='2022' AND HIST_PERIOD='8'")
# print(lw_query)
lw_data=pd.read_sql(lw_query, lawson)
lw_data.to_excel('data_lawson_6.xlsx')
print('finish')
            