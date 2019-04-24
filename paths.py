import os

startPath = os.getcwd()

savePath = "py_Output"
if os.path.exists(savePath) is False:
        os.mkdir(savePath)

os.chdir(savePath)
