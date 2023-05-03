from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait  # available since 2.4.0
from selenium.webdriver.support import expected_conditions as EC  # available since 2.26.0
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options

import pandas as pd

import pyautogui
import time
import csv


class DTEK_REFRESHER:
    def __init__(self):
        self.df = pd.read_csv(r'C:\Users\V Song\Documents\1112_APPROVED.csv', header=None, index_col=None)
        
        self.geckodriver_fpath = r'C:\Users\V Song\Desktop\_PYTHONSCRIPTS\_webdriver\geckodriver.exe'
        
        self.firefox_binary_location = r'C:\Program Files\Mozilla Firefox\firefox.exe'
        self.driver = None

        self.main_sub()
        
    def main_sub(self):
        
        print(self.df)
        

        self.launch_webdriver()
        time.sleep(0.5)
        self.deltek_login()
        time.sleep(0.5)
        
        return
    
    def launch_webdriver(self):

        print('Launching webdriver...\n')
        time.sleep(1)

        # launches firefox
        o = Options()
        s = Service(fr'{self.geckodriver_fpath}')
        
        o.binary_location = self.firefox_binary_location
        self.driver = webdriver.Firefox(service=s, options=o)
        time.sleep(4)

        # changes active tab to newly opened tab (index 1)
        self.driver.switch_to.window(self.driver.window_handles[1])

        return

    def deltek_login(self):

        # GET url
        self.driver.get('https://url.com')

        print('Logging into Deltek PORTAL...')
        time.sleep(2)
        # get DOM element using xpath to get username textbox

        # self.driver.execute_script('window.open("")')
        self.driver.switch_to.window(self.driver.window_handles[1])

        # login page: emailfield and passw and the textfield for login information
        user_name = self.driver.find_element(By.XPATH, '//*[@id="userID"]')
        user_name.send_keys('USERNAME')

        passw = self.driver.find_element(By.XPATH, '//*[@id="password"]')
        passw.send_keys('PASSWORD')

        time.sleep(3)

        self.driver.find_element(By.ID, 'loginBtn').click()
        time.sleep(5)
        
        elementpage = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[3]/div[2]/div/div[3]/div/div[1]/div[3]/div[1]/div[3]/span[3]/span')))
        
        
        for proj in self.df[self.df.columns[0]]:
            time.sleep(2)
            self.driver.find_element(By.CLASS_NAME, 'dropdown-field').send_keys(proj)
            time.sleep(1)
            
            # self.driver.find_element(By.XPATH, '/html/body/div[2]/div[3]/div[2]/div/div[3]/div/div[1]/div[3]/div[1]/div[3]/span[7]').click()
            time.sleep(1)
            # self.driver.find_element(By.XPATH, '/html/body/div[22]/div/ul/li[1]/div/a/div/div').click()
            time.sleep(1)
            print(f'refreshing proj {proj}')
            print('doing something..')

    
            pyautogui.moveTo(1500, 234)
            time.sleep(1)
            pyautogui.click()
            time.sleep(2)
            pyautogui.moveTo(1500, 273)
            time.sleep(1)
            pyautogui.click()
            time.sleep(3)    

def main():
    dr = DTEK_REFRESHER()
    
if __name__ == '__main__':
    main()
    