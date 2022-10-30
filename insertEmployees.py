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

excel_data_df = pd.read_excel('toUpload.xlsx')
excel_data_df = excel_data_df.fillna("")
df = excel_data_df.reset_index()  # make sure indexes pair with number of rows

for index, row in df.iterrows():
    FirstName = row['FirstName']
    MiddleName = row['MiddleName']
    LastName = row['LastName']
    OfficeId = row['OfficeId']
    id=engine.execute("INSERT INTO Employees(FirstName, MiddleName, LastName, OfficeId)VALUES (?, ?, ?,?)",(FirstName,MiddleName,LastName,OfficeId))
    #print(row['FirstName'],row['MiddleName'],row['LastName'],row['OfficeId'])
    #print(type(OfficeId))
    


#id=engine.execute("INSERT INTO Employees(FirstName, MiddleName, LastName, OfficeId)VALUES ('Cardinal', 'Stavanger', 'Norway',12)")

