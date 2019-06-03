import os
from datetime import datetime

today = datetime.now().date()
savePath = today.strftime("%Y%m%d")

startPath = os.getcwd()

# Check if directory exists
        # !Exists, Create folder without int
        # Exists, create new with int appended

def create_output_folder():
        num = 1
        savePath = today.strftime("%Y%m%d")
        # Check for a folder named todays date
        while os.path.exists(savePath) is True:
                # Name folder with number appended to the date, increment number, then check again
                print('Todays date folder exists')
                savePath = today.strftime("%Y%m%d")
                savePath = savePath + "-" + str(num)
                num += 1
        # Once there is no folder with same name, create a new one and change to that directory
        if os.path.exists(savePath) is False:
                print("Create new folder named " + str(savePath))
                os.mkdir(savePath)
        os.chdir(savePath)

