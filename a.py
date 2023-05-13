import os
from win32com.client import Dispatch
import wget
import zipfile

def get_version_via_com(filename):
    parser = Dispatch("Scripting.FileSystemObject")
    version = parser.GetFileVersion(filename)
    return version

edgepath = r"C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
if os.path.exists(edgepath):
    edgeversion = get_version_via_com(edgepath)
    print("Your msedge version is ",edgeversion)
    driverpath = r"./MicrosoftWebDriver.exe"
    if os.path.exists(driverpath):
        driverversion = get_version_via_com(driverpath)
        print("Your msedgeDriver version is ",driverversion)
        if driverversion == edgeversion:
            edge_prepared = 1
        else:
            #wget.download("https://msedgedriver.azureedge.net/" + edgeversion + "/edgedriver_win64.zip")
            file = zipfile.ZipFile("./edgedriver_win64.zip")
            file.extractall("./")
            file.close()
            os.rename("./msedgedriver.exe", "./MicrosoftWebDriver.exe")
    else:
        #wget.download("https://msedgedriver.azureedge.net/" + edgeversion + "/edgedriver_win64.zip")
        file = zipfile.ZipFile("./edgedriver_win64.zip")
        file.extractall("./")
        file.close()
        os.rename("./msedgedriver.exe", "./MicrosoftWebDriver.exe")