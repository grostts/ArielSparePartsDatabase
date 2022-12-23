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
    driver.implicitly_wait(10)

    # entering a password
    password_input = driver.find_element(By.ID,'Input_Password')
    password_input.clear()
    password_input.send_keys(ariel_password)
    driver.implicitly_wait(10)

    # press login button
    login_button_click = driver.find_element(By.XPATH, '//button[contains(@class, "")]').click()
    driver.implicitly_wait(20)
    time.sleep(10)


    # Open ariel parts page
    driver.get(url='https://www.arielcorp.com/parts/portal/')
    driver.implicitly_wait(10)

    # press equipment button
    equipment_button = driver.find_element(By.ID, 'tab-1449-btnInnerEl')
    equipment_button.click()
    driver.implicitly_wait(10)

    print('Start downloading the bom files...')
    counter = 0
    for elem in serial_numbers:
        # fill in the field with the serial number
        text_field = driver.find_element(By.ID, 'textfield-1181-inputEl')
        text_field.clear()
        text_field.send_keys(elem)

        # press search button
        search_button = driver.find_element(By.ID, 'button-1186')
        search_button.click()
        driver.implicitly_wait(10)

        try:
            # press parts button
            parts_button = driver.find_element(By.ID, 'button-1202-btnEl')
            parts_button.click()
            driver.implicitly_wait(10)
            try:
                # press view as list button
                view_as_list_button = driver.find_element(By.ID, 'button-1246-btnEl')
                view_as_list_button.click()
                driver.implicitly_wait(10)

                # downloading BOM file
                download_button = driver.find_element(By.ID, 'button-1241-btnInnerEl')
                download_button.click()
                driver.implicitly_wait(10)
            except:
                # downloading BOM file
                download_button = driver.find_element(By.ID, 'button-1241-btnInnerEl')
                download_button.click()
                driver.implicitly_wait(10)

            # back to the equipment page
            back = driver.find_element(By.ID, 'button-1288-btnInnerEl')
            back.click()
            driver.implicitly_wait(30)

            counter += 1
            current_time = time.time()

            print(f'{counter}/ {total_sn} was downloaded. {round(current_time - start_time, 1)} seconds have passed.')

        finally:
            continue
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
