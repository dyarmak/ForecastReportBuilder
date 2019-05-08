import datetime

today = datetime.datetime.now().date()
year = today.strftime("%Y")
month = today.strftime("%m")
day = today.strftime("%d")

# FileName variables
forecastFName = "qry_Forecast.xlsx"
invoicedFName = "qry_Invoiced.xlsx"
creditFName = "qry_Credits.xlsx"
combinedFName = "Combined.xlsx"
unformattedFName = "Unformatted.xlsx"
outputFName = "Forecast" + year + month + day + ".xlsx"

# Set to 1 to created logging files
logMe = 0