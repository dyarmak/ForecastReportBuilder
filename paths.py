import os
from datetime import datetime

today = datetime.now().date()
savePath = today.strftime("%Y%m%d")

startPath = os.getcwd()

if os.path.exists(savePath) is False:
        os.mkdir(savePath)

os.chdir(savePath)
