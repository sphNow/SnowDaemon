
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from SNDaemon.config import *
from web.my_web_browser import my_web_browser

# RUTAS
url = 'https://santander.service-now.com/nav_to.do?uri=%2Fhome.do'
path_chromedriver = '../chromedriver_win32/chromedriver.exe'  # Ruta del Chromedriver

class my_snow:
    chrome_browser = None

    bannerXPathClass = 'none'

    snow_new_id = "//a[@id='requesturl']"
    #Iframe
    snow_frame_id = "gsft_main"
    snow_frame_name = "iframe"

    # Req Fields
    snow_send = "//button[@id='oi_order_now_button']"
    snow_add_file = "//button[@id='sc_attachment_button']"
    snow_add_title = "//input[@id='loadFileXml']"
    snow_add_input = "//input[@id='attachFile']"
    snow_file_list = "//div[@id='header_attachment']"
    snow_file_close = "//button[@id='attachment_closemodal']"

    snow_system_op = "//span//input[@type='checkbox']"

    snow_company_field = "//div [@class='input-group']//input"
    snow_desc_field = "//div//textarea"

    # Issue Fields
    snow_num_ticket = '//input[@id="sys_readonly.incident.number"]'
    snow_company = '//input[@id="sys_display.incident.company"]'
    snow_type = '//select[@id="incident.u_type"]'
    snow_category = '//select[@id="incident.category"]'
    snow_subcategory = 'incident.subcategory'
    snow_environment = '//select[@id="incident.u_used_for"]'
    snow_app = '//input[@id="sys_display.incident.cmdb_ci"]'
    snow_group_assign = '//div[@id="element.incident.assignment_group"]'
    snow_short_desc = '//input[@id="incident.short_description"]'
    snow_desc = '//textarea[@id="incident.description"]'

    # %% Builder/Destructor
    def __init__(self, driverInfo=None, urlPath=None, pathInput=None):
        self.chrome_browser = my_web_browser(driverInfo, urlPath, pathInput)
        self.__doOpenWeb()

    def __del__(self):
        if self.chrome_browser:
            del self.chrome_browser

# %% Class f(x)'s
    def __doOpenWeb(self):
        self.chrome_browser.driver.get(urlSnow)

        if not self.chrome_browser.driver.find_element(By.TAG_NAME, self.snow_frame_name) is None:
            self.chrome_browser.driver.switch_to.frame(
                self.chrome_browser.driver.find_element(By.TAG_NAME, self.snow_frame_name))

        """ wait for log in to pop up """
        WebDriverWait(self.chrome_browser.driver, minWaitT).until(
            EC.element_to_be_clickable((By.XPATH, self.snow_send)))


    def do_Send_Snow(self, paramA, paramB, paramC, file_path):
        env = str(paramC).split("-")
        if len(env) > 1 and env[1] == "MUREX EUROPA": #Murex 2 EUR
            self.__doActionClickByScript(self.snow_system_op, 0)
            self.__doActionClickByScript(self.snow_system_op, 1)
        elif len(env) > 1 and env[1] ==  "MUREX EQUITY":
            self.__doActionClickByScript(self.snow_system_op, 2)
        elif len(env) > 1 and env[1] ==  "MUREX LATAM":
            self.__doActionClickByScript(self.snow_system_op, 3)
        elif len(env) > 1 and env[1] ==  "Murex Brasil 3":
            self.__doActionClickByScript(self.snow_system_op, 4)

        self.__doInputByXPATH_Elements(self.snow_company_field, -1, "SCIB Global", self.snow_send)
        self.__doInputByXPATH_Elements(self.snow_desc_field, 0, "Email Sailpoint Request (" + str(paramA) + ")\n" + str(paramC)+ ")\n" + str(paramB), self.snow_send)

        self.__doClickByXPATH_Elements(self.snow_add_file, 0, self.snow_add_title)
        self.__doDragAndDropXPATH_Elements(self.snow_add_input, file_path, 0, "//a[text()='" + str(file_path).split("/")[-1] + "']")

        self.__doClickByXPATH_Elements(self.snow_file_close, 0, self.snow_send)
        self.__doClickByXPATH_Elements(self.snow_send, 0, self.snow_new_id)
        return self.__doReadByXPATH_Elements(self.snow_new_id, 0)


# %% Snow f(x)'s
    def __doReadByXPATH_Elements(self, xpath, pos):
        value = ''
        webElements = self.chrome_browser.driver.find_elements(By.XPATH, xpath)
        if len(webElements) > pos:
            value = webElements[pos].get_attribute('innerText')
        return value

    def __doWaitLoadByXPath(self, element=bannerXPathClass):
        webElements = self.chrome_browser.driver.find_elements(By.XPATH, element)

        while len(webElements) > 0 and webElements[0].get_attribute("style") == self.bannerXPathClass:
            webElements = self.chrome_browser.driver.find_elements(By.XPATH, element)

    def __doInputByXPATH_Elements(self, xpath, pos, val, wait_xpath=None):
        try:
            if wait_xpath:
                WebDriverWait(self.chrome_browser.driver, minWaitT).until(
                    EC.element_to_be_clickable((By.XPATH, wait_xpath)))
            webElements = self.chrome_browser.driver.find_elements(By.XPATH, xpath)
            if len(webElements) > pos:
                webElements[pos].send_keys(val)
        except Exception as e:
            print(e)
        finally:
            pass

    def __doActionClickByScript(self, xpath, pos, waited_xpath=None):
        webElements = self.chrome_browser.driver.find_elements(By.XPATH, xpath)
        if len(webElements) > pos:
            self.chrome_browser.driver.execute_script("arguments[0].click();", webElements[pos])
            if waited_xpath:
                WebDriverWait(self.chrome_browser.driver, minWaitT).until(
                    EC.element_to_be_clickable((By.XPATH, waited_xpath)))

    def __doClickByXPATH_Elements(self, xpath, pos, wait_xpath=None):
        self.__doWaitLoadByXPath(self.bannerXPathClass)
        webElements = self.chrome_browser.driver.find_elements(By.XPATH, xpath)
        if len(webElements) > pos:
            webElements[pos].click()
            if wait_xpath:
                WebDriverWait(self.chrome_browser.driver, minWaitT).until(
                    EC.element_to_be_clickable((By.XPATH, wait_xpath)))

    def __doDragAndDropXPATH_Elements(self, xpath, filePath, position, wait_xpath=None):
        try:
            webElements = self.chrome_browser.driver.find_elements(By.XPATH, xpath)
            if len(webElements) > position:
                webElements[position].send_keys(filePath)
                if wait_xpath:
                    WebDriverWait(self.chrome_browser.driver, minWaitT).until(
                        EC.element_to_be_clickable((By.XPATH, wait_xpath)))
        except Exception as e:
            print("Error DragAndDrop.\n"+str(e))
        finally:
            pass

    def __doActionClickBy_XPATH_Elements(self, xpath, pos, wait_xpath=None):
        self.__doWaitLoadByXPath(self.bannerXPathClass)
        webElements = self.chrome_browser.driver.find_elements(By.XPATH, xpath)
        if len(webElements) > pos:
            ActionChains(self.chrome_browser.driver).move_to_element(webElements[pos]).click().perform()
            if wait_xpath:
                WebDriverWait(self.chrome_browser.driver, minWaitT).until(
                    EC.element_to_be_clickable((By.XPATH, wait_xpath)))