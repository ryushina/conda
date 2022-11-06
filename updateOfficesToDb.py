import pandas as pd
import pyodbc
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
table_df = pd.read_sql_table(
    'Offices',
    con=engine,
    columns=[
        'Id',
        'OffcAcr',
        'OfficeName'
    ]
)

df = pd.read_csv('Offices.csv')


#print(table_df)

for index, row in df.iterrows():
    for index1, row1 in table_df.iterrows():
       if row['OfficeAcr'] == row1['OffcAcr']:
             df.loc[index,'OfficeId'] = row1['Id']
             df.loc[index,'OfficeName'] = row1['OfficeName']


df.to_csv('OfficeRef.csv',index=False)