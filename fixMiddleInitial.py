import pandas as pd
import pyodbc
import numpy as np
import sqlalchemy as sa
from sqlalchemy import create_engine, MetaData,Table, Column, Numeric, Integer, VARCHAR
from sqlalchemy.exc import SQLAlchemyError
from sqlalchemy.engine import result
import urllib
import openpyxl

params = urllib.parse.quote_plus("DRIVER={SQL Server Native Client 11.0};"
                                 "SERVER=DESKTOP-T7S1ENJ\SQL2014;"
                                 "DATABASE=ICT;" 
                                 "Trusted_Connection=yes")

engine = sa.create_engine("mssql+pyodbc:///?odbc_connect={}".format(params))

df = pd.read_csv('Employees.csv')
#find MiddleName that possibly don't have period
short = df[df["MiddleName"].str.len() < 2 ]
#put period
for index, row in short.iterrows():
    short.loc[index,'MiddleName'] = short.loc[index,'MiddleName'] + "."
#replace dataframe middlename with replaced values
for index1, row1 in df.iterrows():
    for index2, row2 in short.iterrows():
        if row1['Id'] == row2['Id']:
            df.loc[index1,'MiddleName'] = short.loc[index2,'MiddleName']
#test if there periods are inserted
noperiods = df[df["MiddleName"].str.len() < 2 ]
#find MiddleName that = to .,"",NaN
NaNMiddleName = df[df['MiddleName'].isna()]
#fill nanmiddlename with ''
for index, row in NaNMiddleName.iterrows():
    df.loc[index,'MiddleName'] = ''
#check if all that middleName has no periods
short2 = df[df["MiddleName"].str.len() == 1 ]
midnan2 = df[df['MiddleName'].isna()]
if short2.empty:
    print("Your dataframe is successfully cleared and with periods on middle initial")
if midnan2.empty:
    print("Your dataframe is successfully cleared and with no nulls")
#find other irregular data
irre = df[df["MiddleName"].str.len() > 2 OR ]
#alldata_balance = alldata[(alldata[IBRD] !=0) | (alldata[IMF] !=0)]