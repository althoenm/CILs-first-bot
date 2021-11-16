# %%
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from secrets import username, password, test_record
import getpass
import time

# %%
# Open Chrome
driver = webdriver.Chrome('C:/Users/matte/Downloads/chromedriver_win32/chromedriver.exe')

# Open the website
driver.get("https://www.cilsfirst.com/")

# Wait for the page to load
time.sleep(2)

# %%
# Ask for username and password
# try:
#     user = input("Username: ")
#     p = getpass.getpass("Password: ")
# except:
#     print("Error: username or password not recognized")
#     driver.close()

user = username
p = password
    

# %%
# Find the username and password fields and enter the credentials
try:
    username_path = driver.find_element_by_xpath('/html/body/form/center/div/table/tbody/tr[1]/td[2]/input').send_keys(user)
    password_path = driver.find_element_by_xpath('/html/body/form/center/div/table/tbody/tr[2]/td[2]/input').send_keys(p)
except:
    print("Could not find the username and password fields")
# %%
# Find the login button and click it
submit = driver.find_element_by_xpath('/html/body/form/center/div/table/tbody/tr[3]/td/input').click()
time.sleep(2)

# %%
# Find the Individual Services button and click it

i_path = driver.find_element_by_xpath('//*[@id="side_2"]/a').click()
time.sleep(2)
# %%
# Find 'add new individual service record' button and click it
add_new_path = driver.find_element_by_xpath('//*[@id="content_div"]/div[2]/table/tbody/tr[4]/td/a').click()
# %%
# Identify the fields and enter the data
contact_date = driver.find_element_by_xpath('//*[@id="in_118clone"]').send_keys(test_record['contact_date'])
time.sleep(3)

# %%
time_begun = driver.find_element_by_xpath('//*[@id="in_528"]').send_keys(test_record['time_begun'])
time.sleep(.25)

# %%
time_ended = driver.find_element_by_xpath('//*[@id="in_529"]').send_keys(test_record['time_ended'])
time.sleep(.25)

# %%
hours = driver.find_element_by_xpath('//*[@id="in_123"]').send_keys(test_record['hours'])
time.sleep(.25)
# %%
consumer_field = driver.find_element_by_xpath('//*[@id="in_119"]')
time.sleep(1.5)

# %%
individual_path = driver.find_element_by_xpath(f"//*[contains(text(), '{test_record['consumer']}')]").click()
time.sleep(.25)

# %%
service_path = driver.find_element_by_xpath('//*[@id="in_120"]').send_keys(test_record['service'])
time.sleep(.25)
# %%
priority_area_path = driver.find_element_by_xpath('//*[@id="in_340"]').send_keys(test_record['priority_area'])
time.sleep(.25)
# %%
funding_source = driver.find_element_by_xpath('//*[@id="in_127"]').send_keys(test_record['funding_source'])
time.sleep(.25)
# %% 
staff_comments = driver.find_element_by_xpath('//*[@id="in_132"]').send_keys(test_record['staff_coments'])
time.sleep(.25)

# %%
ir = driver.find_element_by_xpath('//*[@id="in_135"]').send_keys(test_record['ir'])
