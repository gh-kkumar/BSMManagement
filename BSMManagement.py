import time
import openpyxl
import xlsxwriter
import WebElementReusability as WER
import ReadWriteDataFromExcel as RWDE
import BrowserElementProperties as BEP
#import OrderManagement as OM
#import os
#import win32com.client as comclt
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from pathlib import Path
from selenium.webdriver.common import keys


FilePath = str(Path().resolve()) + r'\Excel Files\UrlsForProject.xlsx'
Sheet = 'Portal Urls'
ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Url')
Url = str(RWDE.ReadData(FilePath, Sheet, 5, ColumnNo))

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('disable-notifications')
driver = webdriver.Chrome(executable_path = str(Path().resolve()) + '\Browser\chromedriver_win32\chromedriver', options=chrome_options)
driver.maximize_window()
driver.get(Url)
#print(driver.title)

FilePath = str(Path().resolve()) + '\Excel Files\BSMManagement.xlsx'
Sheet = 'BSM Page Data'
RowCount = RWDE.RowCount(FilePath, Sheet)
Seconds = 1

#1. This is for SFDC Login

for RowIndex in range(2, RowCount + 1):
    ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Run')
    if(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo) == 'Y'):
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.CSS_SELECTOR, '#username', 60)
        ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'User Name')
        Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo))

        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.CSS_SELECTOR, '#password', 60)
        ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'Password')
        Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo))

        time.sleep(Seconds)
        ColumnNo = RWDE.FindColumnNoByName(FilePath, Sheet, 'RememberMe')
        if(RWDE.ReadData(FilePath, Sheet, RowIndex, ColumnNo) == 'Y'):
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.CSS_SELECTOR, '#rememberUn', 60)
            Element.click()

        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.CSS_SELECTOR, '#Login', 60)
        Element.click()

        FilePath1 = str(Path().resolve()) + '\Excel Files\PhlebotomistManagement.xlsx'
        Sheet1 = 'Phlebotomist page Data'
        ColumnNo1 = RWDE.FindColumnNoByName1(FilePath1, Sheet1, 'BCK ID', 5)

        RowNo1 = 6
        RowCount1 = RowNo1 + 2
        for RowNo in range(RowNo1, RowCount1):
            #ColumnNo2 = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Run', 5)
            #RWDE.WriteData(FilePath, Sheet, RowNo, ColumnNo2, 'Y')

            ColumnNo2 = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Scan Barcode', 5)
            BCKID = RWDE.ReadData(FilePath1, Sheet1, RowNo, ColumnNo1)
            RWDE.WriteData(FilePath, Sheet, RowNo, ColumnNo2, BCKID)

            ColumnNo2 = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube1 Secondary Check', 5)
            RWDE.WriteData(FilePath, Sheet, RowNo, ColumnNo2, 'Y')

            ColumnNo2 = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube2 Secondary Check', 5)
            RWDE.WriteData(FilePath, Sheet, RowNo, ColumnNo2, 'Y')

            ColumnNo2 = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube3 Secondary Check', 5)
            RWDE.WriteData(FilePath, Sheet, RowNo, ColumnNo2, 'Y')

            ColumnNo2 = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube4 Secondary Check', 5)
            RWDE.WriteData(FilePath, Sheet, RowNo, ColumnNo2, 'Y')

        for Rowindex1 in range(6, RowCount + 1):
            ColumnNo = RWDE.FindColumnNoByName1(FilePath1, Sheet1, 'Run', 5)
            if(RWDE.ReadData(FilePath1, Sheet1, Rowindex1, ColumnNo) == 'Y'):
                # BSM Button Div
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[. = "BSM"]', 60)
                driver.execute_script('arguments[0].click();', Element)

                # Scan BarCode TextBox
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "Scan Barcode"]', 60)
                ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Scan Barcode', 5)
                if (str(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo)) != 'None'):
                    Element.send_keys(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo))
                else:
                    Element.send_keys('')

                # Scan BCK
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[. ="Scan BCK"]', 60)
                driver.execute_script('arguments[0].click();', Element)

                # Tube1 TextBox
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[2]/div[2]/div[1]/div[1]/div/lightning-input//input', 60)
                ColumnNo = RWDE.FindColumnNoByName1(FilePath1, Sheet1, 'Tube1', 5)
                if (str(RWDE.ReadData(FilePath1, Sheet1, Rowindex1, ColumnNo)) != 'None'):
                    ColumnNo = RWDE.FindColumnNoByName1(FilePath1, Sheet1, 'Tube1', 5)
                    Element.send_keys(RWDE.ReadData(FilePath1, Sheet1, Rowindex1, ColumnNo))

                    Tube1 = RWDE.ReadData(FilePath1, Sheet1, Rowindex1, ColumnNo)
                    ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube1', 5)
                    RWDE.WriteData(FilePath, Sheet, Rowindex1, ColumnNo, Tube1)
                else:
                    Element.send_keys('')

                ##Tube 1 Volume in ml
                #ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube1 Volume(in ml)', 5)
                #if(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo) != 'None'):
                #    time.sleep(Seconds)
                #    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[1]/div[2]/div/lightning-input/div/input', 60)
                #    Element.send_keys(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo))

                ## Tube1 Exception
                #ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube1 Exception', 5)
                #if (RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo) != 'None'):
                #    time.sleep(Seconds)
                #    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[1]/div[3]//input[@placeholder = "Select an Option"]', 60)
                #    driver.execute_script('arguments[0].click();', Element)

                #    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[1]/div[3]//span[2][. = "' + RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo) + '"]', 60)
                #    driver.execute_script('arguments[0].click();', Element)

                # Tube1 Secondary Check
                ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube1 Secondary Check', 5)
                if (str(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo)) == 'Y'):
                    time.sleep(Seconds)
                    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[1]/div[4]/lightning-input/div/span//span[1]', 60)
                    Element.click()

                # Tube2 TextBox
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[2]/div[1]/div/lightning-input/div/input', 60)
                ColumnNo = RWDE.FindColumnNoByName1(FilePath1, Sheet1, 'Tube2', 5)
                if (str(RWDE.ReadData(FilePath1, Sheet1, Rowindex1, ColumnNo)) != 'None'):
                    ColumnNo = RWDE.FindColumnNoByName1(FilePath1, Sheet1, 'Tube2', 5)
                    Element.send_keys(RWDE.ReadData(FilePath1, Sheet1, Rowindex1, ColumnNo))

                    Tube2 = RWDE.ReadData(FilePath1, Sheet1, Rowindex1, ColumnNo)
                    ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube2', 5)
                    RWDE.WriteData(FilePath, Sheet, Rowindex1, ColumnNo, Tube2)
                else:
                    Element.send_keys('')

                ## Tube 2 Volume in ml
                #ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube2 Volume(in ml)', 5)
                #if (RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo) != 'None'):
                #    time.sleep(Seconds)
                #    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[2]/div[2]/div/lightning-input/div/input', 60)
                #    Element.send_keys(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo))

                ## Tube2 Exception
                #ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube2 Exception', 5)
                #if (RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo) != 'None'):
                #    time.sleep(Seconds)
                #    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[2]/div[3]//input[@placeholder = "Select an Option"]', 60)
                #    driver.execute_script('arguments[0].click();', Element)

                #    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[2]/div[3]//span[2][. = "' + RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo) + '"]', 60)
                #    driver.execute_script('arguments[0].click();', Element)

                # Tube2 Secondary Check
                time.sleep(Seconds)
                ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube1 Secondary Check', 5)
                if (str(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo)) == 'Y'):
                    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[2]/div[4]/lightning-input/div/span//span[1]', 60)
                    Element.click()

                # Tube3 TextBox
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[3]/div[1]/div/lightning-input//input', 60)
                ColumnNo = RWDE.FindColumnNoByName1(FilePath1, Sheet1, 'Tube3', 5)
                if (str(RWDE.ReadData(FilePath1, Sheet1, Rowindex1, ColumnNo)) != 'None'):
                    ColumnNo = RWDE.FindColumnNoByName1(FilePath1, Sheet1, 'Tube3', 5)
                    Element.send_keys(RWDE.ReadData(FilePath1, Sheet1, Rowindex1, ColumnNo))

                    Tube3 = RWDE.ReadData(FilePath1, Sheet1, Rowindex1, ColumnNo)
                    ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube3', 5)
                    RWDE.WriteData(FilePath, Sheet, Rowindex1, ColumnNo, Tube3)
                else:
                    Element.send_keys('')

                # Tube3 Secondary Check
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[3]/div[4]/lightning-input/div/span//span[1]', 60)
                ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube3 Secondary Check', 5)
                if (str(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo)) == 'Y'):
                    Element.click()

                # Tube4 TextBox
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[4]/div[1]/div//input', 60)
                ColumnNo = RWDE.FindColumnNoByName1(FilePath1, Sheet1, 'Tube4', 5)
                if (str(RWDE.ReadData(FilePath1, Sheet1, Rowindex1, ColumnNo)) != 'None'):
                    ColumnNo = RWDE.FindColumnNoByName1(FilePath1, Sheet1, 'Tube4', 5)
                    Element.send_keys(RWDE.ReadData(FilePath1, Sheet1, Rowindex1, ColumnNo))

                    Tube4 = RWDE.ReadData(FilePath1, Sheet1, Rowindex1, ColumnNo)
                    ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube4', 5)
                    RWDE.WriteData(FilePath, Sheet, Rowindex1, ColumnNo, Tube4)
                else:
                    Element.send_keys('')

                # Tube4 Secondary Check
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[4]/div[4]/lightning-input/div/span//span[1]', 60)
                ColumnNo = RWDE.FindColumnNoByName1(FilePath, Sheet, 'Tube4 Secondary Check', 5)
                if (str(RWDE.ReadData(FilePath, Sheet, Rowindex1, ColumnNo)) == 'Y'):
                    Element.click()

                # Receive Button
                time.sleep(Seconds)
                Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[. = "Receive"]', 60)
                driver.execute_script('arguments[0].click();', Element)

                if(Rowindex1 == RowCount):
                    # Account Icon
                    time.sleep(3)
                    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button/div/span//span', 60)
                    driver.execute_script('arguments[0].click();', Element)

                    # Log Out Link
                    time.sleep(Seconds)
                    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//a[. = "Log Out"]', 60)
                    driver.execute_script('arguments[0].click();', Element)

                    time.sleep(2)
                    driver.quit()
