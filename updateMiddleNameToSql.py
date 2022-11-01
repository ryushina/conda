import pandas as pd
import pyodbc
import sqlalchemy as sa
from sqlalchemy import create_engine, MetaData,Table, Column, Numeric, Integer, VARCHAR, text
from sqlalchemy.exc import SQLAlchemyError
from sqlalchemy.engine import result
import urllib
import openpyxl
from pandas import *

params = urllib.parse.quote_plus("DRIVER={SQL Server Native Client 11.0};"
                                 "SERVER=DESKTOP-T7S1ENJ\SQL2014;"
                                 "DATABASE=ICT;" 
                                 "Trusted_Connection=yes")

engine = sa.create_engine("mssql+pyodbc:///?odbc_connect={}".format(params))

df = pd.read_csv('fixedMiddleInitial.csv')


sql_text = text("""
    UPDATE Employees
    SET MiddleName = :middlename
    WHERE Id = :id
""")


    
    
table_df = pd.read_sql_table(
    'Employees',
    con=engine,
    columns=[
        'FirstName',
        'LastName',
    ]
)


for index, row in df.iterrows():
    if pd.isnull(df.loc[index,'MiddleName']):
        df.loc[index,'MiddleName'] = ''
    with engine.connect() as conn:
        result = conn.execute(sql_text, middlename=df.loc[index,'MiddleName'], id=row['Id'])


    
