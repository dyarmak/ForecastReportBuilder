# Backup the SQL Query to a date stamped folder.


import os
import shutil
from os import path
import datetime
from excelFNames import forecastFName, invoicedFName, creditFName

today = str(datetime.datetime.now().date())
rawSQLPath = "rawQry-" + today
if os.path.exists(rawSQLPath) is False:
        os.mkdir(rawSQLPath)

# File names
rawForecastQry = today + "-qry_Forecast_raw.xlsx"
rawInvoicedQry = today + "-qry_Invoiced_raw.xlsx"
rawCreditsQry = today + "-qry_Credits_raw.xlsx"


if os.path.exists(forecastFName):
    src = os.path.realpath(forecastFName)
    head, tail = path.split(src)
    # print("Path: " + head)
    # print("File: " + tail)
    dstFolder = head + "\\" + rawSQLPath
    print("Dst folder: " + dstFolder)
    tail = rawForecastQry
    dst = dstFolder + "\\" + tail
    shutil.copy(src, dst)

if os.path.exists(invoicedFName):
    src = os.path.realpath(invoicedFName)
    head, tail = path.split(src)
    # print("Path: " + head)
    # print("File: " + tail)
    dstFolder = head + "\\" + rawSQLPath
    print("Dst folder: " + dstFolder)
    tail = rawInvoicedQry
    dst = dstFolder + "\\" + tail
    shutil.copy(src, dst)

if os.path.exists(creditFName):
    src = os.path.realpath(creditFName)
    head, tail = path.split(src)
    # print("Path: " + head)
    # print("File: " + tail)
    dstFolder = head + "\\" + rawSQLPath
    print("Dst folder: " + dstFolder)
    tail = rawCreditsQry
    dst = dstFolder + "\\" + tail
    shutil.copy(src, dst)
print("")