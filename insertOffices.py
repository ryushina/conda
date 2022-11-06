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


sql_text = text("""
    INSERT INTO Offices(OffcAcr, OfficeName)VALUES (:offcacr, :officename)
""")


with engine.connect() as conn:
    result1 = conn.execute(sql_text, offcacr="GK", officename="Gawad Kalinga")
    result2 = conn.execute(sql_text, offcacr="EU", officename="Municipal Health Offices")
    result3 = conn.execute(sql_text, offcacr="GO-NVWMC", officename="GO-Watershed")

