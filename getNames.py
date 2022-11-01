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

df = pd.read_excel('Properties.xlsx')

lastname = []
firstname = []
middlename = []

for index, row in df.iterrows():
    mn = row['Employee'].split(",")
    lastname.append(mn[0])
    ln_ready = mn[1]
    ln_and_mn = ln_ready.split(" ")
    if len(ln_and_mn) == 5 and '.(' in ln_and_mn[2]:
        firstname.append(ln_and_mn[1])
        middlename.append(ln_and_mn[2][0:2])
    elif len(ln_and_mn) == 5 and 'II' in ln_and_mn[3]:
        firstname.append(ln_and_mn[1]+" "+ln_and_mn[3])
        middlename.append(ln_and_mn[2])
    elif len(ln_and_mn) == 3 and len(ln_and_mn[1]) > 2 and len(ln_and_mn[2]) > 2:
        firstname.append(ln_and_mn[1]+" "+ln_and_mn[2])
        middlename.append("")
    elif len(ln_and_mn) == 3:
        firstname.append(ln_and_mn[1])
        middlename.append(ln_and_mn[2])
    elif len(ln_and_mn) == 4 and ')' in ln_and_mn[3]:
        firstname.append(ln_and_mn[1])
        middlename.append(ln_and_mn[2])
    elif len(ln_and_mn) == 4 and '.' in ln_and_mn[3]:
        firstname.append(ln_and_mn[1] + ' '+ln_and_mn[2])
        middlename.append(ln_and_mn[3])
    elif len(ln_and_mn) == 2:
        firstname.append(ln_and_mn[1])
        middlename.append("")
    elif len(ln_and_mn) == 5 and ')' in ln_and_mn[4]:
        firstname.append(ln_and_mn[1] + ' '+ln_and_mn[2])
        middlename.append(ln_and_mn[3])
    elif len(ln_and_mn) == 5 and '.' in ln_and_mn[4]:
        firstname.append(ln_and_mn[1] + ' '+ln_and_mn[2]+" "+ln_and_mn[2])
        middlename.append(ln_and_mn[4])
    elif len(ln_and_mn) == 4 and len(ln_and_mn[3])==1:
        firstname.append(ln_and_mn[1]+ ' '+ln_and_mn[2])
        middlename.append(ln_and_mn[3]+".")
    elif len(ln_and_mn) == 4 and '.' in ln_and_mn[2]:
        firstname.append(ln_and_mn[1]+" "+ln_and_mn[3])
        middlename.append(ln_and_mn[2])
    elif len(ln_and_mn) == 4 and len(ln_and_mn[3]) > 2 and len(ln_and_mn[2]) > 2:
        firstname.append(ln_and_mn[1]+" "+ln_and_mn[2])
        middlename.append(ln_and_mn[3])
    else:
        firstname.append("unknown")
        middlename.append("unknown")
        
df['First Name'] = firstname
df['Middle Name'] = middlename
df['Last Name'] = lastname
df.to_csv('names_extracted.csv',index=False)

####TOOLS
# elm_count = firstname.count('unknown')
# print(elm_count)

# firstname.index("unknown")

# print(df.loc[[159220]])
# ln_and_mn = df.loc[129,"Employee"].split(",")
# lnn = ln_and_mn[1].split(" ")    