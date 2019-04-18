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

print("*** WARNING ***")
print("This script will query the DB. This may take a few minutes.")
input("Press Enter to continue")

startTimer = time.time()

savePath = "py_Output"
if os.path.exists(savePath) is False:
        os.mkdir(savePath)

os.chdir(savePath)

# RUN SQLtoExcel
import SQLtoExcel
# RUN Forecast manipulation code
import ForecastManipulations
# RUN Credits manipulation code
import CreditsManipulations
# RUN Invoiced manipulation code
import InvoicedManipulations
# Amalgamate the three files into one
import joinData
# Add VLookup values
import vlook


# -------------- Timer -----------------------
endTimer = time.time()
buildTime = endTimer-startTimer
print("Report building took: " + str(buildTime) + "seconds\n")

input('Press Enter to close:')





