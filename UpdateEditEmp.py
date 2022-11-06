#script for updating middlename of employees based from generated "fixMiddleInitial.csv"
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


sql_update = text("""
    UPDATE Employees
    SET FirstName = :firstname,
    MiddleName = :middlename,
    LastName = :lastname
    WHERE Id = :id
""")

sql_insert = text("""
    INSERT INTO Employees (FirstName, MiddleName, LastName,OfficeId)
    VALUES (:firstname, :middlename, :lastname, :officeid);
""")


df = pd.read_csv('employeesToAddCsv.csv')
table_df = pd.read_sql_table(
    'Employees',
    con=engine,
    columns=[
        'Id',
        'FirstName',
        'MiddleName',
        'LastName',
        'OfficeId',
        'Position'
    ]
)


for index, row in df.iterrows():
    withMatch = False
    matchedId = 0
    for index1, row1 in table_df.iterrows():
        if row['First Name'].upper() == row1['FirstName'].upper() and row['Last Name'].upper() == row1['LastName'].upper():
            withMatch = True
            matchedId = row1['Id']
    if withMatch == True:
        with engine.connect() as conn:
            result = conn.execute(sql_update, firstname=df.loc[index,'First Name'].upper(), middlename=df.loc[index,'Middle Name'].upper(),lastname=df.loc[index,'Last Name'].upper(), id=int(matchedId))
            print("updated "+str(index))
    else:
        with engine.connect() as conn:
            result = conn.execute(sql_insert, firstname=str(df.loc[index,'First Name']).upper(), middlename=str(df.loc[index,'Middle Name']).upper(),lastname=str(df.loc[index,'Last Name']).upper(), officeid=int(df.loc[index,'OfficeId']))
            print("inserted "+str(index))
        



