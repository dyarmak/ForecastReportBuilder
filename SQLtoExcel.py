#*********************************************************************************
# This script will query the DB and save the query results into excel files
#*********************************************************************************

## SQL Query export to Excel files

import os
import pandas as pd
import pyodbc
from queries import forecastQuery, invoicedQuery, creditQuery
from excelFNames import forecastFName, invoicedFName, creditFName

# Define connection variables 
server = os.environ.get('DBServer')
database = os.environ.get('DBName')
username = os.environ.get('DBUsername')
password = os.environ.get("DBPassword")
# Establish DB Connection
cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

# Execute Queries
foreDF = pd.read_sql(forecastQuery, cnxn, index_col='SubProjectID')
invoDF = pd.read_sql(invoicedQuery, cnxn, index_col='SubProjectID')
credDF = pd.read_sql(creditQuery, cnxn, index_col='SubProjectID')


# export to .xlsx
foreDF.to_excel(forecastFName, sheet_name='Sheet1')
invoDF.to_excel(invoicedFName, sheet_name='Sheet1')
credDF.to_excel(creditFName, sheet_name='Sheet1')

print("Data from 3 queries pulled from DB and exported to .xlsx files\n")