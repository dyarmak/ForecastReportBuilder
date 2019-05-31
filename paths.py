import os
from datetime import datetime

today = datetime.now().date()
savePath = today.strftime("%Y%m%d")

startPath = os.getcwd()

# Check if directory exists
        # !Exists, Create folder without int
        # Exists, create new with int appended
num = 1
while os.path.exists(savePath) is True:
        savePath = today.strftime("%Y%m%d")
        savePath = savePath + "-" + str(num)
        num += 1

if os.path.exists(savePath) is False:
        os.mkdir(savePath)

os.chdir(savePath)
