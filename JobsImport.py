from asyncio.windows_events import NULL
import psycopg2 as db_connect
import os
from datetime import datetime
import pandas as pd
from cProfile import label
from lib2to3.pgen2 import driver
from typing import OrderedDict
from unittest import result
from PyQt6.QtWidgets import QApplication, QHBoxLayout, QProgressBar, QVBoxLayout, QWidget,QLabel, QLineEdit, QPushButton,QListWidget, QMessageBox
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import Qt
import os
import time
from PyQt6.QtCore import QThread, pyqtSignal
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from os.path import exists
import json
from selenium.webdriver.support.ui import Select
from time import sleep

driver = webdriver.Chrome(executable_path="C:/Users/roshan.liu/Scripts/SAP_DailyRelease/chromedriver.exe")
driver.get("https://partners.gorenje.com/sagcc/Sredina.aspx")
wait = WebDriverWait(driver, 10)
with open('config.json','r') as j:
    config= json.load(j)
    usr= config['usr']
    pwd = config['pwd']
driver.find_element(By.ID, "usr").send_keys(usr)
driver.find_element(By.ID, "pwd").send_keys(pwd)
driver.find_element(By.ID, "btnPrijava").click()
driver.get("https://partners.gorenje.com/sagcc/potni_pregled1.aspx")
Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_drpCenter")).select_by_value('0')
driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_dtDatumFilter_B-1").click()
wait.until(EC.visibility_of_element_located((By.ID,"ctl00_ContentPlaceHolder1_dtDatumFilter_DDD_C_BC")))
driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_dtDatumFilter_DDD_C_BC").click()
driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnPrikazi").click()
driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lnkBCsv").click()

while True:
    if os.path.isfile('C://Users//roshan.liu//Downloads//ASPxGridView1.csv'):
        break
    else:
        time.sleep(10)
print('file downloaded')
