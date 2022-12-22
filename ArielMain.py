import time
from selenium import webdriver
from selenium.webdriver.common.by import By
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

# set webdriver options
chromeOptions = webdriver.ChromeOptions()
prefs = {"download.default_directory" : "C:\\Users\\hp\\GitHubProjects\\ArielSparePartsDatabase\\Downloads"}
chromeOptions.add_experimental_option("prefs",prefs)
chromeOptions.add_argument("--disable-blink-features=AutomationControlled")
# headless mode
# chromeOptions.headless = True
driver = webdriver.Chrome(options=chromeOptions)


try:
    print('Login Ariel Website...')
    # Open and login  Ariel website
    driver.get(url=url)

    # entering a email
    email_input = driver.find_element(By.ID,'Input_UserName')
    email_input.clear()
    email_input.send_keys(ariel_login)
    time.sleep(5)

    # entering a password
    password_input = driver.find_element(By.ID,'Input_Password')
    password_input.clear()
    password_input.send_keys(ariel_password)
    time.sleep(5)

    # press login button
    login_button_click = driver.find_element(By.XPATH, '//button[contains(@class, "")]').click()
    time.sleep(10)

except Exception as ex:
    print(ex)
finally:
    driver.close()
    driver.quit()


# Program runtime output
end_time = time.time()
elapsed_time = end_time - start_time
print(f'Program running time  = {round(elapsed_time)} seconds')