from SNDaemon.config import *

from selenium import webdriver

import os

class my_web_browser:
    driverInfo = ''
    urlPath = ''
    pathInput = ''
    driver = None

    # %% Builder/Destructor
    def __init__(self, driverInfo=None, urlPath=None, pathInput=None):
        self.driverInfo = driverInfo if driverInfo else driver_path
        self.urlPath = urlPath if urlPath else url_main
        self.pathInput = pathInput if pathInput else dwnload_path
        self.__prepare_dwnload_folder()
        self.__initChrome()

    def __del__(self):
        try:
            if self.driver:
                self.driver.close()
                self.driver.quit()
        except:
            pass

    # %% Class f(x)'s
    def __initChrome(self):
        # Open up Chrome Webbrowser
        chromeOptions = webdriver.ChromeOptions()
        chromeOptions.add_argument("--start-maximized")
        prefs = {"profile.default_content_settings.popups": 0,
                 "download.default_directory": self.pathInput,
                 "download.directory_upgrade": True,
                 "directory_upgrade": True}
        chromeOptions.add_experimental_option("prefs", prefs)
        self.driver = webdriver.Chrome(executable_path=self.driverInfo, options=chromeOptions)

    def __prepare_dwnload_folder(self):
        try:
            if not os.path.exists(self.pathInput): os.makedirs(self.pathInput, exist_ok=True)
        except Exception as e:
            print("Error trying to create download folder. \n\n" + str(e))