#*********************************************************************************
# This script will Build the forecast from Scratch
# 1) Query DB -> Save Forecast, Invoiced and Credits to excel files
# 2) Manipulate those excel files
# 3) join them
# 4) perform the required vlook up for setting 'Type'
#*********************************************************************************

import time
import os
from paths import create_output_folder
from excelFNames import unformattedFName

startTimer = time.time()

create_output_folder()

print('**Running ForecastBuilder.py')
import ForecastBuilder

# Format final report
print("**Running ForecastFormatter.py")
import ForecastFormatter

# -------------- Timer -----------------------
endTimer = time.time()
buildTime = endTimer-startTimer
print("Forecast Report Building took: " + str(buildTime) + "seconds\n")

os.remove(unformattedFName)



