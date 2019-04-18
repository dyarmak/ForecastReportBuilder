#*********************************************************************************
# This script will query the DB and save the query results into excel files
#*********************************************************************************


## From SQL to DataFrame Pandas
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
# All of those need to be put into environ Vars before I upload to GitHub
cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

credDF = pd.read_sql(creditQuery, cnxn, index_col='SubProjectID')
# export to .xlsx
credDF.to_excel(creditFName, sheet_name='Sheet1')

invoDF = pd.read_sql(invoicedQuery, cnxn, index_col='SubProjectID')
# export to .xlsx
invoDF.to_excel(invoicedFName, sheet_name='Sheet1')

foreDF = pd.read_sql(forecastQuery, cnxn, index_col='SubProjectID')
# export to .xlsx
foreDF.to_excel(forecastFName, sheet_name='Sheet1')


