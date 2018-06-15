import pandas as pd
from selenium import webdriver
import time
import glob
import os

# define variables
doc = 'PATH/FILENAME.XLS OF DOC'
driver = webdriver.Chrome(executable_path='PATH TO CHROMEDRIVER.EXE')
sheet1 = 'NAME OF SHEET1'
sheet2 = 'NAME OF SHEET2'
fb_url = 'https://business.facebook.com/login'
fb_username = 'EMAIL ADDRESS'
fb_password = 'PASSWORD'

# create dataframe of all data in each sheet
df1 = pd.read_excel(io=doc, sheet_name=sheet1, header=0) # insert number of header rows to skip
df2 = pd.read_excel(io=doc, sheet_name=sheet2, header=0) # insert number of header rows to skip

# create list of hyperlinks from first column (index=0) of each dataframe
links1 = df1.values.T[0].tolist()
links2 = df2.values.T[0].tolist()

# login to Facebook
driver.get(fb_url)
driver.find_element_by_id("email").send_keys(fb_username)
driver.find_element_by_id("pass").send_keys(fb_password)
driver.find_element_by_id("loginbutton").click()

# loop through hyperlinks in each list (these particular reports save to /downloads)
for url in links1:
    urls = "'" + url + "'"
    print(urls, '\n')
    driver.get(url)
    time.sleep(3)

for url in links2:
    urls = "'" + url + "'"
    print(urls, '\n')
    driver.get(url)
    time.sleep(3)










