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
table_df = pd.read_sql_table(
    'PPEs',
    con=engine,
    columns=[
        'Id',
        'EmployeeId',
        'PropertyNo',
        'Description',
        'PurchaseDate',
        'Article',
        'Remarks',
        'OfficeAcr'
    ]
)

df = pd.read_csv('names_extracted.csv')


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

for index, row in df.iterrows():
    PropertyNum = 0
    EmployeeId = 0
    for index1, row1 in table_df.iterrows():
        if row['fldPropertyNo'] == row1['PropertyNo']:
            PropertNum = row['fldPropertyNo']
    if PropertyNum == 0:
        print("No match")
    else:
        print(f"Match {PropertyNum}")
 
def getEmployeeId(firstname,lastname):
  return 5 * x


