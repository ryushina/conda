import pandas as pd
import pyodbc
import sqlalchemy as sa
from sqlalchemy import create_engine, MetaData,Table, Column, Numeric, Integer, VARCHAR
from sqlalchemy.exc import SQLAlchemyError
from sqlalchemy.engine import result
import urllib
import openpyxl

# params = urllib.parse.quote_plus("DRIVER={SQL Server Native Client 11.0};"
#                                  "SERVER=DESKTOP-T7S1ENJ\SQL2014;"
#                                  "DATABASE=ICT;" 
#                                  "Trusted_Connection=yes")

# engine = sa.create_engine("mssql+pyodbc:///?odbc_connect={}".format(params))

df = pd.read_csv('names_extracted_manual.csv')

print(df)

office_id = []
offices = df.fldOfficeID.unique().tolist()


# for index, row in df.iterrows():
#     if row['fldOfficeID'] == 'PPDO':
#         office_id.append(41)
#     else:

####TOOLS
# elm_count = firstname.count('unknown')
# print(elm_count)

# firstname.index("unknown")

# print(df.loc[[159220]])
# ln_and_mn = df.loc[129,"Employee"].split(",")
# lnn = ln_and_mn[1].split(" ")   