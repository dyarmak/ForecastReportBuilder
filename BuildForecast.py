#*********************************************************************************
# This script will Build the forecast from Scratch
# 1) Query DB -> Save Forecast, Invoiced and Credits to excel files
# 2) Manipulate those excel files
# 3) join them
# 4) perform the required vlook up for setting 'Type'
#*********************************************************************************
print("*** WARNING ***")
print("This script will execute SELECT queries on the DB:\n 1) qry_forecast\n 2) qry_invoiced\n 3) qry_credits")
print("It will then combine them to build the details sheet of the ForecastReport.")
print("It will then build the Summary, byPM, ByClient, and byType pages of the Forecast Report.")
print("This may take a couple minutes to execute.")
input("Press Enter if you wish to continue, or close this window.")
print("")

import os
from os import path
import time
from excelFNames import forecastFName, invoicedFName, creditFName, combinedFName
from paths import create_output_folder

startTimer = time.time()

# create_output_folder()

# # Query DB and save to excel
# print("**Running SQLtoExcel.py")
# import SQLtoExcel
# # Save raw SQL to folder
# print("**Backing up raw query results")
# import saveRawSQL
# # Align Headings
# print("**Running alignHeadings.py")
# import alignHeadings
# # Set default DUE Values
# print("**Running setDueValues.py")
# import setDueValues
# # RUN Forecast manipulation code
# print("**Running ForecastManipulations.py")
# import ForecastManipulations
# # RUN Credits manipulation code
# print("**Running CreditsManipulations.py")
# import CreditsManipulations
# # RUN Invoiced manipulation code
# print("**Running InvoiceManipulations.py")
# import InvoicedManipulations
# # Amalgamate the three files into one
# print("**Running joinData.py")
# import joinData
# # Add VLookup values
print("**Running vlook.py")
import vlook
# Build report tabs
print("**Running buildTabs.py")
import buildTabs
# Format final report
print("**Running excelFormatting.py")
import excelFormatting

# -------------- Timer -----------------------
endTimer = time.time()
buildTime = endTimer-startTimer
print("Forecast Report Building took: " + str(buildTime) + "seconds\n")

input('Press Enter to close:')





