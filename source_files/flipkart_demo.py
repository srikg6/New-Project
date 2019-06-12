from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import os
from source_files import locators
import xlwt
import unittest
import time

# Global variables
file_path = 'C:\\Users\\Srikanth\\Desktop\\example.xls'
# Username and password are set as enviroment variables
username = os.getenv('gmail_usrnm')
password = os.getenv('gmail_pwd')

URL = "https://www.flipkart.com/"
Sheetname = 'testdata'
myacc = "My Account"
fail_message = "Login Failed"


class FlipkartDemo(unittest.TestCase):
    driver = webdriver.Chrome()
    driver.get(URL)
    driver.maximize_window()

    def test_findallthelinks(self):
        try:
            un = self.driver.find_element_by_xpath(locators.username_txt)
            un.send_keys(username)
            pw = self.driver.find_element_by_xpath(locators.password_txt)
            pw.send_keys(password)
            login_btn = self.driver.find_element_by_xpath(locators.login_btn)
            login_btn.click()
            time.sleep(3)
            wait = WebDriverWait(self.driver, 5)
            #  Wait for 5 seconds before throwing an exception if element is
            # not found
            my_account = wait.until(EC.visibility_of_element_located((
                By.XPATH, locators.my_account)))
            assert my_account.text == myacc, fail_message
            search_box = self.driver.find_element_by_class_name(
                    locators.search_txt)
            search_box.send_keys("iphone", Keys.ENTER)
            wb = xlwt.Workbook()
            sheet1 = wb.add_sheet(Sheetname)
            Apple_iPhone_6s_RoseGold_32GB_model = wait.until(
                EC.visibility_of_element_located(
                    (By.XPATH, locators.Apple_iPhone_6s_Rose_Gold_32GB_model)))
            sheet1.write(0, 0, Apple_iPhone_6s_RoseGold_32GB_model.text)
            Apple_iPhone_6s_RoseGold_32GB_price = self.driver.find_element_by_xpath(
                locators.Apple_iPhone_6s_Rose_Gold_32GB_price)
            sheet1.write(0, 1, Apple_iPhone_6s_RoseGold_32GB_price.text)
            Apple_iPhone_X_Space_Gray_64GB_model = self.driver.find_element_by_xpath(
                locators.Apple_iPhone_X_Space_Gray_64GB_model)
            sheet1.write(1, 0, Apple_iPhone_X_Space_Gray_64GB_model.text)
            Apple_iPhone_X_Space_Gray_64GB_price = self.driver.find_element_by_xpath(
                locators.Apple_iPhone_X_Space_Gray_64GB_price)
            sheet1.write(1, 1, Apple_iPhone_X_Space_Gray_64GB_price.text)
            wb.save(file_path)
        except Exception as ex:
            print(ex)
        finally:
            self.driver.close()


if __name__ == '__main__':
    unittest.main()
