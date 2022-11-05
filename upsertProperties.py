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

#test commit and push


