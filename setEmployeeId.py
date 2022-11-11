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

df = pd.read_csv('names_extracted.csv')
df["EmployeeId"] = 0

df_Employees = pd.read_sql_table('Employees', engine)



# for index, row in df.iterrows():
#     for index1, row1 in df_Employees.iterrows():
#         if row['First Name'] == row1['FirstName'] and row['Last Name'] == row1['LastName']:
#             df.loc[index,'EmployeeId'] = row1['Id']
            
# # for index, row in df.iterrows():
# #     employeeId = 0
# #     for index1, row1 in df_Employees.iterrows():
# #         if employeeId != 0:
# #             break
# #         if row1['FirstName'].upper() in row['Employee'].upper() and row1['LastName'].upper() in row['Employee'].upper():
# #             employeeId = row1['Id']
# #     df.loc[index,'EmployeeId'] = employeeId
# #     print(df.loc[[index]])

# df.to_csv('GeneratedEmpId.csv',index=False)

df2 = df.loc[(df['Last Name'] == 'BEJARIN')]
df3 = df_Employees.loc[(df_Employees['LastName'] == 'BEJARIN')]
#28
#1639

print(df2[["Employee"]])
print("------")
print(df3[["FirstName","LastName"]])
print("RESULT OF FIND:")
firstname = df3.loc[1639,'FirstName']
print(str(firstname))

if 'MONICA' in df2.loc[28,'Employee'] :
    print('MONICA is found in dataframe 2 through text')
else:
    print('Not found')

if firstname in df2.loc[28,'Employee'] :
    print('MONICA is found in dataframe through variable')
else:
    print('Not found through variable')



