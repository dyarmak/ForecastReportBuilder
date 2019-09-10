# This script will be run as a Scheduled Task
# It Performs the following:
# - Query Database for Forecast, Invoices, and Credits in the current year
# - Cleans, and manipulates that data
# - Outputs the unformatted data to 'Unformatted.xlsx'

# Setup and load data into DataFrames
# Imports
import os
import pandas as pd
import numpy as np
import pyodbc
import openpyxl
from datetime import datetime
from excelFNames import unformattedFName

# Read from .sql files
def readScriptFromFile(filename):
    # Open and read the file as a single buffer
    fd = open(filename, 'r')
    sqlFile = fd.read()
    fd.close()
    return sqlFile

forecastQuery = readScriptFromFile(r'C:\Users\dyarmak\Documents\SQL Server Management Studio\Queries\Forecast Report\ForecastQuery.sql')
invoicedQuery = readScriptFromFile(r'C:\Users\dyarmak\Documents\SQL Server Management Studio\Queries\Forecast Report\InvoicedQuery.sql')
creditQuery = readScriptFromFile(r'C:\Users\dyarmak\Documents\SQL Server Management Studio\Queries\Forecast Report\CreditsQuery.sql')
profitCenterQuery = readScriptFromFile(r'C:\Users\dyarmak\Documents\SQL Server Management Studio\Queries\Forecast Report\ProfitCenterVlookQuery.sql')
groupsQuery = readScriptFromFile(r'C:\Users\dyarmak\Documents\SQL Server Management Studio\Queries\Forecast Report\SubProjectGroupsQuery.sql')

# Define connection variables 
server = os.environ.get('DBServer')
database = os.environ.get('DBName')
username = os.environ.get('DBUsername')
password = os.environ.get("DBPassword")

# Establish DB Connection
cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

# Execute the queries and store in DFs
foreDF = pd.read_sql(forecastQuery, cnxn, index_col='SubProjectID')
invoDF = pd.read_sql(invoicedQuery, cnxn, index_col='SubProjectID')
credDF = pd.read_sql(creditQuery, cnxn, index_col='SubProjectID')
profitCenter = pd.read_sql(profitCenterQuery, cnxn)
groups = pd.read_sql(groupsQuery, cnxn)


# ## Write Raw Query Results to folder
# Save Raw SQL Query results
print('Backing up raw query results')

today = datetime.now().strftime('%Y%m%d')
# File names
rawForecastQry = today + "-qry_Forecast_raw.xlsx"
rawInvoicedQry = today + "-qry_Invoiced_raw.xlsx"
rawCreditsQry = today + "-qry_Credits_raw.xlsx"


rawSQLPath = "rawQry-" + today
if os.path.exists(rawSQLPath) is False:
        os.mkdir(rawSQLPath)

startDir = os.getcwd()
os.chdir(startDir + '\\' + rawSQLPath)

foreDF.to_excel(rawForecastQry)
invoDF.to_excel(rawInvoicedQry)
credDF.to_excel(rawCreditsQry)

# Back to starting Directory
os.chdir(startDir)

print("Performing Data Manipulations")

# Format Floats to 2 decimal places
pd.set_option('display.float_format', lambda x: '%.2f' % x)

# foreDF
foreDF['Forecast'] = pd.to_numeric(foreDF['Forecast'])
foreDF['Quoted'] = pd.to_numeric(foreDF['Quoted'])
foreDF['OriginalForecast'] = pd.to_numeric(foreDF['OriginalForecast'])
foreDF['Budget'] = pd.to_numeric(foreDF['Budget'])
foreDF['SubTotal'] = pd.to_numeric(foreDF['SubTotal'])

# invoDF
invoDF['OriginalSubTotal'] = pd.to_numeric(invoDF['OriginalSubTotal'])
invoDF['Quoted'] = pd.to_numeric(invoDF['Quoted'])
invoDF['OriginalForecast'] = pd.to_numeric(invoDF['OriginalForecast'])
invoDF['Budget'] = pd.to_numeric(invoDF['Budget'])
invoDF['SubTotal'] = pd.to_numeric(invoDF['SubTotal'])

# credDF
credDF['CreditAmt'] = pd.to_numeric(credDF['CreditAmt'])
credDF['Quoted'] = pd.to_numeric(credDF['Quoted'])
credDF['OriginalForecast'] = pd.to_numeric(credDF['OriginalForecast'])
credDF['Budget'] = pd.to_numeric(credDF['Budget'])
credDF['SubTotal'] = pd.to_numeric(credDF['SubTotal'])

# ## Replace all 0s with np.nan (not a number)
# Because I will be using logic to check for nan ( isna() and notna() funcs ), I need to remove the 0s and replace with np.nan
# 
# Only do this for columns I'll be doing tests on
# - Forecast
# - Quoted
# - OriginalForecast
# - Budget
# - SubTotal

# Forecast DataFrame
foreDF.loc[foreDF.Forecast==0, ['Forecast']]= np.nan
foreDF.loc[foreDF.Quoted==0, ['Quoted']]= np.nan
foreDF.loc[foreDF.OriginalForecast==0, ['OriginalForecast']]= np.nan
foreDF.loc[foreDF.Budget==0, ['Budget']]= np.nan
foreDF.loc[foreDF.SubTotal==0, ['SubTotal']]= np.nan

# invoDF
invoDF.loc[invoDF.OriginalSubTotal==0, ['OriginalSubTotal']]= np.nan
invoDF.loc[invoDF.Quoted==0, ['Quoted']]= np.nan
invoDF.loc[invoDF.OriginalForecast==0, ['OriginalForecast']]= np.nan
invoDF.loc[invoDF.Budget==0, ['Budget']]= np.nan
invoDF.loc[invoDF.SubTotal==0, ['SubTotal']]= np.nan

# credDF
credDF.loc[credDF.CreditAmt==0, ['CreditAmt']]= np.nan
credDF.loc[credDF.Quoted==0, ['Quoted']]= np.nan
credDF.loc[credDF.OriginalForecast==0, ['OriginalForecast']]= np.nan
credDF.loc[credDF.Budget==0, ['Budget']]= np.nan
credDF.loc[credDF.SubTotal==0, ['SubTotal']]= np.nan

foreDF['Due'] = foreDF['Due Date'].dt.strftime('%m-%Y')
invoDF['Due'] = 'Act' + invoDF['Due Date'].dt.strftime('%m-%Y')
credDF['Due'] = 'Act' + credDF['Due Date'].dt.strftime('%m-%Y')

# The columns are now aligned, and Due values are set appropriately


# ## Set displayOrder
# We need to be able to sort on something other than Due, since PowerBI wants to sort if alphabetically.
# 
# Might need to just use numbers to order each column for the report viz

# #### This needs to be re-worked to sort properly
foreDF['displayOrder'] = foreDF['Due Date'].dt.strftime('%Y-%m')
invoDF['displayOrder'] = invoDF['Due Date'].dt.strftime('%Y-%m') + 'Act'
credDF['displayOrder'] = credDF['Due Date'].dt.strftime('%Y-%m') + 'Act'

# # Profit Center - from profitCenter
# Use the profitCenter DataFrame to set the 'ProfitCenter' value FOR EACH SubProjectID - FOR EACH DF

# # set_ProfitCenter_col() definition

def set_ProfitCenter_col(inputDF,lookup):
    df = inputDF.copy()
    # Get column list in proper order
    cols = df.columns.tolist()
    cols.append('ProfitCenter')
    # Reset Index
    df = df.reset_index()
    # Merge profitCenter to the DataFrame
    df = pd.merge(df, lookup, on='SubProjectTypeID')
    
    # Set mask values
    swMask = df.MasterProjectName == 'Software'
    hseMask = df.MasterProjectName == 'HSE'
    crbMask = df.MasterProjectName == 'CRB'
    # Set Profit Center of EDS and MARS to Software
    edsMask = df.ProfitCenter=='EDS'
    marsMask = df.ProfitCenter=='MARS'

    # Use mask to select rows and apply changes
    df.loc[swMask, 'ProfitCenter'] = 'Software'    
    df.loc[edsMask, 'ProfitCenter'] = 'Software'
    df.loc[marsMask, 'ProfitCenter'] = 'Software'
    df.loc[hseMask, 'ProfitCenter'] = 'HSE'
    df.loc[crbMask, 'ProfitCenter'] = 'CRB'

    
    # Set index on SubProjectID
    df = df.set_index('SubProjectID')
    
    # Re=order the columns
    df = df[cols]
    return df

foreDF = set_ProfitCenter_col(foreDF, profitCenter)
invoDF = set_ProfitCenter_col(invoDF, profitCenter)
credDF = set_ProfitCenter_col(credDF, profitCenter)

# # Type - from Groups
# This comes from tbl_SubProjectType and tbl_SubProjectType_Group
# Grab just the columns we need from groups query
groups = groups[['SubProjectTypeID', 'GroupName', 'SortOrder']]

# Must reset_index() before merge, then set_index('SubProjectID') after merge
foreDF = foreDF.reset_index()
invoDF = invoDF.reset_index()
credDF = credDF.reset_index()

foreDF = pd.merge(foreDF, groups, how ='left', on='SubProjectTypeID')
invoDF = pd.merge(invoDF, groups, how ='left', on='SubProjectTypeID')
credDF = pd.merge(credDF, groups, how ='left', on='SubProjectTypeID')


foreDF.rename({'GroupName':'Type'}, axis=1, inplace=True)
invoDF.rename({'GroupName':'Type'}, axis=1, inplace=True)
credDF.rename({'GroupName':'Type'}, axis=1, inplace=True)

foreDF = foreDF.set_index('SubProjectID')
invoDF = invoDF.set_index('SubProjectID')
credDF = credDF.set_index('SubProjectID')


# At this point, the columns are aligned and the Due and Type Columns are set
# With the exception of:
# - foreDF has 'Forecast'
# - invoDF has OriginalSubTotal
# - credDF has 'CreditAmt'

# # Forecast Manipulations

# ## Drop any duplicate SubProjectIDs
# - This comes from a Software SubProject that had 17 interim invoices, blowing the forecast for that project up by 17x

# Find list of duplicated SubProjectIDs
foreDF = foreDF.reset_index()
dup = foreDF.loc[foreDF.duplicated(subset='SubProjectID')]
dupSPs = dup.SubProjectID.unique().tolist()
# sum up the SubTotal Column for a duplicated SubProject
for sp in dupSPs:
    total = 0
    total = foreDF.loc[foreDF.SubProjectID == 24078, 'SubTotal'].sum()
    foreDF.loc[foreDF.SubProjectID == 24078, ['SubTotal']] = total
# remove duplicated SubProjects, keeping the first row.
foreDF = foreDF.drop_duplicates(subset='SubProjectID', keep='first')
foreDF = foreDF.set_index('SubProjectID')


# ## Interim invoicing
# -	In the forecast we have to check if there are invoices against the subproject – I have it in the query to look for invoiced amounts and invoice date sent 
# 
# 
# IF a subproject is in ['Planning', 'In Progress', 'Complete', 'Review'] AND there is an invoiced amt and invoiced date sent
# - Set Forecast = 0 for those subprojects 
# - Otherwise the $ value would show up as an actual in the Invoiced query AND in the Forecast. 
# - There can be an invoiced amount without a date sent – those we can ignore as they have been written but not finalized so we do not use that number.

invoDF = invoDF.reset_index()
# Save list of Interim Invoiced SubProjects
interimInvoicedDF = invoDF.loc[invoDF.duplicated(subset='SubProjectID', keep=False)]
invoDF = invoDF.set_index('SubProjectID')

# list of SubProjectStatus'
status = ['Planning', 'In Progress', 'Complete', 'Review']
statusMask = foreDF.SubProjectStatus.isin(status)

# Filter WHERE 
invoicedMask = foreDF.SubTotal.notna()
invoiceDateMask = foreDF.InvoiceDateSent.notna()

# List of interim invoiced projects still in forecast
interimForecastDF = foreDF.loc[(statusMask & invoicedMask & invoiceDateMask)]

# Set Forecast value on Interim Invoiced SPs to np.nan
foreDF.loc[(statusMask & invoicedMask & invoiceDateMask), ['Forecast']] = np.nan


# ## Find and Delete WHERE Status == Planning AND inBudget == False AND SubTotal == 0

inPlanningMask = foreDF.SubProjectStatus == 'Planning'
notInBudgetMask = foreDF.IncludeinBudget == False
subTotalIsZero = foreDF.SubTotal.isna()
# These are the rows to be dropped
planNotinBudgetSubZeroDF = foreDF.loc[(inPlanningMask & notInBudgetMask & subTotalIsZero), ['ProjectManager', 'ClientName', 'Forecast', 'SubProjectStatus', 'IncludeinBudget', 'SubTotal']]
# Save the rows that are NOT in the selection set
foreDF = foreDF.loc[~(inPlanningMask & notInBudgetMask & subTotalIsZero)].copy()


# ## Drop rows WHERE status == 'Cancel Requested'
# Save list og Cancel Requested SubProjects
cancelRequestedDF = foreDF.loc[(foreDF.SubProjectStatus == 'Cancel Requested')]

# Save the rows WHERE status != 'Cancel Requested'
foreDF = foreDF.loc[~(foreDF.SubProjectStatus == 'Cancel Requested')].copy()

# Check that they were removed
# foreDF.loc[foreDF.SubProjectStatus=='Cancel Requested']


# ## Update Forecast Amounts
# - If invoice has been sent, Set Forecast = np.nan
# - If InvoiceDateSent = None & Quoted != None
# - If invoiceDateSent == None AND Quoted == None AND OriginalForecast != None, Set Forecast = OriginalForecast
# - If invoiceDateSent == None AND Quoted == None AND OriginalForecast == None AND Budget != None, Set Forecast = Budget

# If invoice has been sent, Set Forecast = np.nan
invoicedMask = foreDF.InvoiceDateSent.notna()
foreDF.loc[invoicedMask, ['Forecast']] = np.nan

# If InvoiceDateSent = None & Quoted != None
noInvoicedMask = foreDF.InvoiceDateSent.isna()
quotedMask = foreDF.Quoted.notna()
foreDF.loc[(noInvoicedMask & quotedMask), 'Forecast'] = foreDF.Quoted

# If invoiceDateSent == None AND Quoted == None AND OriginalForecast != None, Set Forecast = OriginalForecast
noInvoicedMask = foreDF.InvoiceDateSent.isna()
noQuotedMask = foreDF.Quoted.isna()
origForeMask = foreDF.OriginalForecast.notna()
foreDF.loc[(noInvoicedMask & noQuotedMask & origForeMask), 'Forecast'] = foreDF.OriginalForecast

# If invoiceDateSent == None AND Quoted == None AND OriginalForecast == None AND Budget != None, Set Forecast = Budget
noInvoicedMask = foreDF.InvoiceDateSent.isna()
noQuotedMask = foreDF.Quoted.isna()
noOrigForeMask = foreDF.OriginalForecast.isna()
budgetMask = foreDF.Budget.notna()
foreDF.loc[(noInvoicedMask & noQuotedMask & noOrigForeMask & budgetMask), 'Forecast'] = foreDF.Budget

# List of projects without a value in Forecast
noForecastValueDF = foreDF.loc[(foreDF.Forecast.isna() & noInvoicedMask)]

# ## Update Due Column
# Mark any overdue items

# If SubProjectStatus == 'Review' or 'Complete'
# Set due to that value
reviewMask = foreDF.SubProjectStatus=='Review'
completeMask = foreDF.SubProjectStatus=='Complete'
foreDF.loc[reviewMask, 'Due'] = foreDF.SubProjectStatus
foreDF.loc[completeMask, 'Due'] = foreDF.SubProjectStatus

# Do the same to displayOrder
foreDF.loc[reviewMask, 'displayOrder'] = foreDF.SubProjectStatus
foreDF.loc[completeMask, 'displayOrder'] = foreDF.SubProjectStatus

todaysDate = datetime.now()

# If subProjectStatus == 'In Progress' or 'Planning' AND dueDate > todaysDate
# Set due to 'Overdue'
inProgressMask = foreDF.SubProjectStatus=='In Progress'
planningMask = foreDF.SubProjectStatus=='Planning'
overDueMask = foreDF['Due Date']<todaysDate

# Set Due = 'Overdue' 
foreDF.loc[(overDueMask & (inProgressMask | planningMask)) & ~(reviewMask | completeMask), 'Due'] = 'Overdue'
foreDF.loc[(overDueMask & (inProgressMask | planningMask)) & ~(reviewMask | completeMask), 'displayOrder'] = 'Overdue'

# Store list of review, completed and overdue
reviewDF = foreDF.loc[reviewMask]
completeDF = foreDF.loc[completeMask]
overDueDF = foreDF.loc[(overDueMask & (inProgressMask | planningMask)) & ~(reviewMask | completeMask)]


# # Credits Manipulations

# All we're doing here is appending 'CR' to each SubProjectID.
# - In order to do that, SubProjectID's Type needs to be converted to string.
# - And in order for the 3 DataFrames to be able to concatenate onto each other later, we'll need to do this to all the DataFrames.
# - Might as well do it here

foreDF = foreDF.reset_index()
invoDF = invoDF.reset_index()
credDF = credDF.reset_index()

foreDF.SubProjectID = foreDF.SubProjectID.apply(str)
invoDF.SubProjectID = invoDF.SubProjectID.apply(str)
credDF.SubProjectID = credDF.SubProjectID.apply(str)

credDF.SubProjectID = credDF.SubProjectID + 'CR'

foreDF = foreDF.set_index('SubProjectID')
invoDF = invoDF.set_index('SubProjectID')
credDF = credDF.set_index('SubProjectID')


# # Invoiced Manipulations
# All we need to do here is check for deferred revenues.

# IF InvoiceDateSent == 2018   Due Gets "Def-1"
hasInvoDateMask = invoDF.InvoiceDateSent.notna()
sentLastYearMask = invoDF.InvoiceDateSent.dt.year==(todaysDate.year - 1)

invoDF.loc[(hasInvoDateMask & sentLastYearMask), 'Due'] = 'Def-1'


# # Combine DataFrames - Join Data

# First we need to rename the CreditAmt and OriginalSubTotal columns to be Forecast

credDF = credDF.rename(columns={'CreditAmt':'Forecast'})
invoDF = invoDF.rename(columns={'OriginalSubTotal':'Forecast'})
combinedDF = foreDF.append(invoDF)
combinedDF = combinedDF.append(credDF)

# # Build Tabs
# In this section the various excel tabs are built out.

# ## Details Tab
detailsDF = combinedDF.copy()


# # Current Year Tabs

# Make list of due columns for the current year
def due_columns():
    """
    Returns a list [] of columns for the "Due" Field
    this would not work on past data...
    It should be modified to somehow use data from the pandas DF.
    """
    today = datetime.now().date()
    curryear = today.strftime("%Y")
    dueCols = ["Def-1"]

    # Act month
    for p in range(1, today.month+1):
        dtString = str(p) + str(curryear)
        dt = datetime.strptime(dtString, '%m%Y').date()
        actString = "Act"+ dt.strftime('%m') + "-" + dt.strftime('%Y')
        dueCols.append(actString)

    # Overdue, completed, Review
    dueCols.append("Overdue")
    dueCols.append("Complete")
    dueCols.append("Review")

    # Future months
    for p in range(today.month, 13):
        dtString = str(p) + str(curryear)
        dt = datetime.strptime(dtString, '%m%Y').date()
        futureString = dt.strftime('%m') + "-" + dt.strftime('%Y')
        dueCols.append(futureString) 

    # Should check the generated list against the a list of unique due values from Pandas?

    return dueCols

dueCols = due_columns()


# ## PM Tab

# grab just the analysis columns
due = combinedDF[["ProjectManager", "ClientName", "Forecast", "Due"]].copy()

# If a PM name is missing / nan, sort() will crash.
# Therefore have to replace with none.
due.fillna({'ProjectManager':'None'}, inplace=True)

# NOW we can pull in list of PMs and sort...

# Get unique PM Names and save to a list
pmNames = due["ProjectManager"].unique().tolist()
pmNames.sort()

# ## Create a new DF grouped by PM, with Due as columns, SumOfForecast as values
sumByPMDF = due.groupby(["ProjectManager", "Due"]).Forecast.sum().unstack()

# The script breaks here if there is an empty month. 
# To fix this we can...
# Create an empty DataFrame with all the columns
blank = pd.DataFrame(columns=dueCols)
blank = blank.append(sumByPMDF, sort=True)

# Set the order of the Due columns as per the list.
sumByPMDF = blank[dueCols]

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


# At this point we have the PM Tab completed.
# We will need to use the data from it to generate the PM Subtotals for Summary TAB


# ## Summary Tab
# - New DataFrame, indexed by PM->Client (a multi-index DataFrame), summed by due period 

pmClient = due.groupby(["ProjectManager", "ClientName", "Due"]).Forecast.sum().unstack()

# The script breaks here if there is an empty month. 
# To fix this we can...
# Create an empty DataFrame with all the columns
blank = pd.DataFrame(columns=dueCols)
blank = pmClient.append(blank, sort=True)
pmClient = blank[dueCols].copy()
idx = pd.MultiIndex.from_tuples(tuples=pmClient.index)
pmClient.index = idx

# Iterate the rows summing Def-1 to 12-2019, storing it in the "Total" column
for r in pmClient.index:
    pmClient.loc[r, "Total"] = pmClient.loc[r, "Def-1":"12-2019"].sum()


# - Here we need to first create a new DF with just the first PMs values
# - Then we can insert the PM column (only be done once) and add their name (for later sorting)
# - Finally we append the PM Subtotal before moving onto the next pm (using the loop)
# - this can definitely be done cleaner, but for now it works great

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


# ## Client Tab

# read in just analysis columns
due = combinedDF[["ClientName", "Forecast", "Due"]].copy()

# Get unique client Names and save to a list
clientNames = due["ClientName"].unique().tolist()
clientNames.sort()

# ## Create a new DF grouped by PM, with Due as columns, SumOfForecast as values
sumByClientDF = due.groupby(["ClientName", "Due"]).Forecast.sum().unstack()

# The script breaks here if there is an empty month. 
# To fix this we can...
# Create an empty DataFrame with all the columns
blank = pd.DataFrame(columns=dueCols)
blank = blank.append(sumByClientDF, sort=True)
sumByClientDF = blank[dueCols].copy()

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


# ## ByType Tab

# read in just analysis columns
typeDF = combinedDF[["Type", "Forecast", "Due"]].copy()

typeDF.fillna({'Type':'None'}, inplace=True)

# Get unique Type Names and save to a list
typeNames = typeDF["Type"].unique().tolist()
typeNames.sort()

# Create a new DF grouped by Type, with Due as columns, SumOfForecast as values
sumByTypeDF = typeDF.groupby(["Type", "Due"]).Forecast.sum().unstack()

# The script breaks here if there is an empty month. 
# To fix this we can...
# Create an empty DataFrame with all the columns
blank = pd.DataFrame(columns=dueCols)
blank = blank.append(sumByTypeDF, sort=True)
sumByTypeDF = blank[dueCols].copy()

# ## Sum each row, store in Totals Column
for t in typeNames:
    sumByTypeDF.loc[t, "Total"] = sumByTypeDF.loc[t, "Def-1":"12-2019"].sum()

# ## Sum Company Wide Total
sumTypeByDue = sumByTypeDF.sum()
companyWideTotals = sumTypeByDue.to_frame(name="Company Wide Total").T
sumByTypeDF = sumByTypeDF.append(companyWideTotals)

# Set index name
sumByTypeDF.index.name = "Type" # Change to Profit Center in future?


# ## By ProfitCenter

# read in just analysis columns
profitCenterDF = combinedDF[["ProfitCenter", "Forecast", "Due"]].copy()

profitCenterDF.fillna({'Type':'None'}, inplace=True)

# Get unique Type Names and save to a list
profitCenterNames = profitCenterDF["ProfitCenter"].unique().tolist()
profitCenterNames.sort()

# Create a new DF grouped by Type, with Due as columns, SumOfForecast as values
sumByProfitCenterDF = profitCenterDF.groupby(["ProfitCenter", "Due"]).Forecast.sum().unstack()

# The script breaks here if there is an empty month. 
# To fix this we can...
# Create an empty DataFrame with all the columns
blank = pd.DataFrame(columns=dueCols)
blank = blank.append(sumByProfitCenterDF, sort=True)
sumByProfitCenterDF = blank[dueCols].copy()

# ## Sum each row, store in Totals Column
for t in profitCenterNames:
    sumByProfitCenterDF.loc[t, "Total"] = sumByProfitCenterDF.loc[t, "Def-1":"12-2019"].sum()

# ## Sum Company Wide Total
sumTypeByDue = sumByProfitCenterDF.sum()
companyWideTotals = sumTypeByDue.to_frame(name="Company Wide Total").T
sumByProfitCenterDF = sumByProfitCenterDF.append(companyWideTotals)

# Set index name
sumByProfitCenterDF.index.name = "ProfitCenter" # Change to Profit Center in future?


# # Next Year Tabs

# Make list of due columns for next year
def next_year_due_columns():
    """Return a list of columns for the "Due" field for NEXT year 
    """
    today = datetime.now().date()
    curryear = today.year
    next_year = datetime(curryear+1, 1, 1)
    dueCols = []

    for p in range(next_year.month, 13):
        dtString = str(p) + str(next_year.year)
        dt = datetime.strptime(dtString, '%m%Y').date()
        futureString = dt.strftime('%m') + "-" + dt.strftime('%Y')
        dueCols.append(futureString) 

    return dueCols

nyDueCols = next_year_due_columns()


# ## NY - PM Tab

# read in just analysis columns
due = combinedDF[["ProjectManager", "ClientName", "Forecast", "Due"]].copy()

# If a PM name is missing / nan, sort() will crash.
# Therefore have to replace with none.
due.fillna({'ProjectManager':'None'}, inplace=True)
# NOW we can pull in list of PMs and sort...

# Get unique PM Names and save to a list
pmNames = due["ProjectManager"].unique().tolist()
pmNames.sort()

# ## Create a new DF grouped by PM, with Due as columns, SumOfForecast as values
ny_sumByPMDF = due.groupby(["ProjectManager", "Due"]).Forecast.sum().unstack().copy()

# The script breaks here if there is an empty month. 
# To fix this we can...
# Create an empty DataFrame with all the columns
ny_blank = pd.DataFrame(columns=nyDueCols)
ny_blank = ny_blank.append(ny_sumByPMDF, sort=True)

# Set the order of the Due columns as per the list.
ny_sumByPMDF = ny_blank[nyDueCols].copy()

# Get list of PM Names from DataFrame
pmNames = ny_sumByPMDF.index.to_list()

# ## Sum by each Due period and save in a Series
ny_sumByDueCol = ny_sumByPMDF.sum()

# ## Transform Series to DataFrame, Transpose, append to sumByPMDF, store in DataFrame
ny_companyWideTotals = ny_sumByDueCol.to_frame(name="Company Wide Total").T
ny_sumByPMDF = ny_sumByPMDF.append(ny_companyWideTotals)

# Set Index Name
ny_sumByPMDF.index.name = "ProjectManager"

# ## Add a totals Column, sum each row, store in Totals Column
for pm in pmNames:
    ny_sumByPMDF.loc[pm, "Total"] = ny_sumByPMDF.loc[pm, :].sum()

# ## Sum Company Wide Grand Total
ny_sumByPMDF.loc["Company Wide Total", "Total"] = ny_sumByPMDF.loc["Company Wide Total"].sum()

# ## Add " Total" to each index in pmSum
new_index = ["{} Total".format(pm) for pm in pmNames]
new_index.append("Company Wide Total")

ny_sumByPMDF.index = new_index


# ## NY Summary Tab

# New DataFrame, indexed by PM->Client (a multi-index DataFrame)
# sum by due period 
ny_pmClient = due.groupby(["ProjectManager", "ClientName", "Due"]).Forecast.sum().unstack()

# The script breaks here if there is an empty month. 
# To fix this we can create an empty DataFrame with all the columns
blank = pd.DataFrame(columns=nyDueCols)
# And append the old original DF to it.
blank = ny_pmClient.append(blank, sort=True)
# THEN copy so we can use the original name again. Not vert slick... but it works
ny_pmClient = blank[nyDueCols].copy()
# Set a multi index
idx = pd.MultiIndex.from_tuples(tuples=ny_pmClient.index)
ny_pmClient.index = idx

# Iterate the rows summing all columns in that row, storing it in the "Total" column
for r in ny_pmClient.index:
    ny_pmClient.loc[r, "Total"] = ny_pmClient.loc[r, :].sum()

# Initialize a new DF with first pm's values
ny_summaryDF = ny_pmClient.loc[pmNames[0]]

# Add ProjectManager column
ny_summaryDF.insert(loc = 0, column="ProjectManager", value=pmNames[0])

# Get first pm Sums
aSum = ny_sumByPMDF.iloc[0]

# cvt to frame and Transpose
aSum = aSum.to_frame(name = ny_sumByPMDF.index[0]).T

# Add ProjectManager column
aSum.insert(loc = 0, column="ProjectManager", value=pmNames[0])

# Append sum to Summary DF
ny_summaryDF = ny_summaryDF.append(aSum, sort=False)

# start loop from 1 i=1
for i in range(1,len(pmNames)):
    client = ny_pmClient.loc[pmNames[i]]
    client.insert(loc = 0, column="ProjectManager", value=pmNames[i])
    aSum = ny_sumByPMDF.iloc[i]
    aSum = aSum.to_frame(name = ny_sumByPMDF.index[i]).T
    aSum.insert(loc = 0, column="ProjectManager", value=pmNames[i])
    ny_summaryDF = ny_summaryDF.append(client, sort=False)
    ny_summaryDF = ny_summaryDF.append(aSum, sort=False)

# Add company wide total, grabbing it from the PM DataFrame
grandTotal = ny_sumByPMDF.loc["Company Wide Total"]
grandTotal = grandTotal.to_frame(name = "Company Wide Total").T
grandTotal.insert(loc = 0, column="ProjectManager", value="H2Safety")
ny_summaryDF = ny_summaryDF.append(grandTotal, sort=False)

# ## GroupBy formatting with multi-index for clean excel output
ny_summaryDF.index.name = "ClientName"
ny_summaryDF.reset_index(drop=False, inplace=True)
ny_summaryDF = ny_summaryDF.set_index(["ProjectManager", "ClientName"])

# export unformatted file to .xlsx
with pd.ExcelWriter(unformattedFName) as writer:
    summaryDF.to_excel(writer, sheet_name='Summary')
    sumByPMDF.to_excel(writer, sheet_name='PM')
    sumByClientDF.to_excel(writer, sheet_name='Client')
    sumByTypeDF.to_excel(writer, sheet_name='Type')
    sumByProfitCenterDF.to_excel(writer, sheet_name='ProfitCenter')
    ny_summaryDF.to_excel(writer, sheet_name='NY-Summary')
    ny_sumByPMDF.to_excel(writer, sheet_name='NY-PM')
    detailsDF.to_excel(writer, sheet_name='Details')

print("Saved " + unformattedFName)















