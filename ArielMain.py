import time
from selenium import webdriver
from AuthorizationData import ariel_password
from AuthorizationData import ariel_login
import pandas as pd


start_time = time.time()

# Get a frame serial number base
file_name = input('Enter the name of file with a serial numbers: ')
result = pd.read_excel(f'{file_name}.xlsx')
serial_numbers = result['sn'].tolist()
total_sn = len(serial_numbers)

# Using Selenium
url = 'https://account.arielcorp.com/Identity/Account/Login'



