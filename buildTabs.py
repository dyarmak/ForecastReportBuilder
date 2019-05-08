# Imports
import os
import pandas as pd
import numpy as np
import openpyxl
from excelFNames import combinedFName, unformattedFName
from myxlutils import due_columns
from paths import savePath # Need to be in the py_output folder

# --------------- Settings --------------- #
# dueCols should come from a variable that can be updated outside of the script, much like the vlook file

# Define the order of the Due Cols that we want to display
# This will need to be changed as we get closer to year end.
# Perhaps there should be a 12 month forecast sheet...?
dueCols = due_columns()

# Formatting floats
pd.set_option('display.float_format', lambda x: '%.2f' % x)

# ******************************************************************* #
# ************************** Details TAB **************************** #
# ******************************************************************* #
# Read in full excel file for Details Sheet and save as details
detailsDF = pd.read_excel(combinedFName, index_col="SubProjectID")

# ---> detailsDF is ready for output

# ******************************************************************* #
# **************************** PM TAB ******************************* #
# ******************************************************************* #

# read in excel with just analysis columns
due = pd.read_excel(combinedFName, usecols=["SubProjectID", "ProjectManager", "ClientName", "Forecast", "Due"])

# Get unique PM Names and save to a list
pmNames = due["ProjectManager"].unique().tolist()
pmNames.sort()
# pmCount = len(pmNames) # Don't think we need this anymore

# ## Create a new DF grouped by PM, with Due as columns, SumOfForecast as values
sumByPMDF = due.groupby(["ProjectManager", "Due"]).Forecast.sum().unstack()

# Set the order of the Due columns as per the list.
sumByPMDF = sumByPMDF[dueCols]

# ## Sum by each Due period and save in a Series
sumByDueCol = sumByPMDF.sum()

# ## Transform Series to DataFrame, Transpose, append to sumByPMDF, store in DataFrame
companyWideTotals = sumByDueCol.to_frame(name="Company Wide Total").T
sumByPMDF = sumByPMDF.append(companyWideTotals)

# Set Index Name
sumByPMDF.index.name = "ProjectManager"

# ## Add a totals Column, sum each row, store in Totals Column
for pm in pmNames:
    sumByPMDF.loc[pm, "Total"] = sumByPMDF.loc[pm, "Def-1":"12-2019"].sum()

# ## Sum Company Wide Grand Total
sumByPMDF.loc["Company Wide Total", "Total"] = sumByPMDF.loc["Company Wide Total"].sum()

# ## Add " Total" to each index in pmSum
new_index = ["{} Total".format(pm) for pm in pmNames]
new_index.append("Company Wide Total")
len(new_index)

sumByPMDF.index = new_index

print("PM Tab created")

# At this point we have the PM Tab completed.
# We will need to use the data from it to generate the PM Subtotals for Summary TAB

# ******************************************************************* #
# ************************ Summary TAB ****************************** #
# ******************************************************************* #

# ## New DataFrame, indexed by PM->Client (a multi-index DataFrame), summed by due period 
pmClient = due.groupby(["ProjectManager", "ClientName", "Due"]).Forecast.sum().unstack()
pmClient = pmClient[dueCols]

# Iterate the rows summing Def-1 to 12-2019, storing it in the "Total" column
for r in pmClient.index:
    pmClient.loc[r, "Total"] = pmClient.loc[r, "Def-1":"12-2019"].sum()

# ## Build Summary-Tab DataFrame

# Here we need to first create a new DF with just the first PMs values
# Then we can insert the PM column (only be done once) and add their name (for later sorting)
# Finally we append the PM Subtotal before moving onto the next pm
# this can definitely be done cleaner, but for now it works great

# Initialize a new DF with first pm's values
summaryDF = pmClient.loc[pmNames[0]]
# Add ProjectManager column
summaryDF.insert(loc = 0, column="ProjectManager", value=pmNames[0])
# Get first pm Sums
aSum = sumByPMDF.iloc[0]
# cvt to frame and Transpose
aSum = aSum.to_frame(name = sumByPMDF.index[0]).T
# Add ProjectManager column
aSum.insert(loc = 0, column="ProjectManager", value=pmNames[0])
# Append sum to Summary DF
summaryDF = summaryDF.append(aSum, sort=False)
# start loop from 1 i=1
for i in range(1,len(pmNames)):
    client = pmClient.loc[pmNames[i]]
    client.insert(loc = 0, column="ProjectManager", value=pmNames[i])
    aSum = sumByPMDF.iloc[i]
    aSum = aSum.to_frame(name = sumByPMDF.index[i]).T
    aSum.insert(loc = 0, column="ProjectManager", value=pmNames[i])
    summaryDF = summaryDF.append(client, sort=False)
    summaryDF = summaryDF.append(aSum, sort=False)


# Add company wide total, grabbing it from the PM DataFrame
grandTotal = sumByPMDF.loc["Company Wide Total"]
grandTotal = grandTotal.to_frame(name = "Company Wide Total").T
grandTotal.insert(loc = 0, column="ProjectManager", value="H2Safety")
summaryDF = summaryDF.append(grandTotal, sort=False)


# ## GroupBy formatting with multi-index for clean excel output
summaryDF.index.name = "ClientName"
summaryDF.reset_index(drop=False, inplace=True)
summaryDF = summaryDF.set_index(["ProjectManager", "ClientName"])

print("Summary Tab created")

# ******************************************************************* #
# ************************** Client TAB **************************** #
# ******************************************************************* #

# read in excel with just analysis columns
due = pd.read_excel(combinedFName, usecols=["SubProjectID", "ClientName", "Forecast", "Due"])

# Get unique client Names and save to a list
clientNames = due["ClientName"].unique().tolist()
clientNames.sort()

# ## Create a new DF grouped by PM, with Due as columns, SumOfForecast as values
sumByClientDF = due.groupby(["ClientName", "Due"]).Forecast.sum().unstack()
sumByClientDF = sumByClientDF[dueCols]

# ## Sum each row, store in Totals Column
for client in clientNames:
    sumByClientDF.loc[client, "Total"] = sumByClientDF.loc[client, "Def-1":"12-2019"].sum()


# ## Sort client by total column
sumByClientDF = sumByClientDF.sort_values(by = "Total", ascending=False)


# ## Sum each Due Column and save to a Series
sumByDueCol = sumByClientDF.sum()


# ## Transform Series to DataFrame, Transpose, append to sumByPM, store in DataFrame
companyWideTotals = sumByDueCol.to_frame(name="Company Wide Total").T
sumByClientDF = sumByClientDF.append(companyWideTotals)

# Set index name
sumByClientDF.index.name = "ClientName"

print("Client Tab created")

# ******************************************************************* #
# ************************** ByType TAB **************************** #
# ******************************************************************* #

# read in excel with just analysis columns
typeDF = pd.read_excel(combinedFName, usecols=["SubProjectID", "Type", "Forecast", "Due"])

# Get unique Type Names and save to a list
typeNames = typeDF["Type"].unique().tolist()
typeNames.sort()

# Create a new DF grouped by Type, with Due as columns, SumOfForecast as values
sumByTypeDF = typeDF.groupby(["Type", "Due"]).Forecast.sum().unstack()
sumByTypeDF = sumByTypeDF[dueCols]

# ## Sum each row, store in Totals Column
for t in typeNames:
    sumByTypeDF.loc[t, "Total"] = sumByTypeDF.loc[t, "Def-1":"12-2019"].sum()

# ## Sum Company Wide Total
sumTypeByDue = sumByTypeDF.sum()
companyWideTotals = sumTypeByDue.to_frame(name="Company Wide Total").T
sumByTypeDF = sumByTypeDF.append(companyWideTotals)

# Set index name
sumByTypeDF.index.name = "Type" # Change to Profit Center in future?

print("Type Tab created")

# ## Export ALL the DataFrames to Excel
# Need to name this with the current date
with pd.ExcelWriter(unformattedFName) as writer:
    summaryDF.to_excel(writer, sheet_name='Summary')
    sumByPMDF.to_excel(writer, sheet_name='PM')
    sumByClientDF.to_excel(writer, sheet_name='ByClient')
    sumByTypeDF.to_excel(writer, sheet_name='ByType')
    detailsDF.to_excel(writer, sheet_name='Details')

print("unformatted tabs saved to " + unformattedFName + "\n")