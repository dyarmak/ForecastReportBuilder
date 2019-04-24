#*********************************************************************************
# This script will Build the forecast from Scratch
# 1) Query DB -> Save Forecast, Invoiced and Credits to excel files
# 2) Manipulate those excel files
# 3) join them
# 4) perform the required vlook up for setting 'Type'
#*********************************************************************************

import os
from os import path
import time
from excelFNames import forecastFName, invoicedFName, creditFName, combinedFName
import paths

print("*** WARNING ***")
print("This script will execute the following SELECT queries on the DB:\nqry_forecast,\nqry_invoiced\nqry_credits")
print("It will then combine them to build the details sheet of the ForecastReport.")
print("This may take a few minutes to execute.")
input("Press Enter to continue")

startTimer = time.time()

# Query DB and save to excel
print("**Running SQLtoExcel.py")
import SQLtoExcel
# Align Headings
print("**Running alignHeadings.py")
import alignHeadings
# Set default DUE Values
print("**Running setDueValues.py")
import setDueValues
# RUN Forecast manipulation code
print("**Running ForecastManipulations.py")
import ForecastManipulations
# RUN Credits manipulation code
print("**Running CreditsManipulations.py")
import CreditsManipulations
# RUN Invoiced manipulation code
print("**Running InvoiceManipulations.py")
import InvoicedManipulations
# Amalgamate the three files into one
print("**Running joinData.py")
import joinData
# Add VLookup values
print("**Running vlook.py")
import vlook


# -------------- Timer -----------------------
endTimer = time.time()
buildTime = endTimer-startTimer
print("Forecast Report Building took: " + str(buildTime) + "seconds\n")

input('Press Enter to close:')





