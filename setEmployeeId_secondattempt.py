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

df = pd.read_csv('GeneratedEmpId.csv')

filtered = df.loc[df['EmployeeId']==0]

df_Employees = pd.read_sql_table('Employees', engine)

EmployeeIds = []
for index, row in df.iterrows():
    firstname = ''
    lastname = ''
    id = 0
    for index1, row1 in df_Employees.iterrows():
        if firstname == '':
            if row1['FirstName'] in row['Employee']:
                firstname = row1['FirstName']
                id = row1['Id']
            else:
                firstname = ''
        if lastname == '':
            if row1['LastName'] in row['Employee']:
                lastname = row1['LastName']
            else:
                lastname = ''
    if firstname != '' and lastname != '':
        EmployeeIds.append(id)
    else:
        EmployeeIds.append(0)  
          

df['completeId'] = EmployeeIds
print(len(EmployeeIds))
df.to_csv('GeneratedEmpIdv2.csv',index=False)


