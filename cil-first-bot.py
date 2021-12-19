# %%
from openpyxl import Workbook, load_workbook
from utils import *
from test.secrets import username, password
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

# %%
# access sheet names
wb = load_workbook('reporting_template.xlsx')

# %%
# Load worksheets by name
tt_sh = wb['time_tracker']
ca_sh = wb['community_activities']

# %%
# Load rows to lists
tt_rows = []
ca_rows =[]
read_excel_rows_to_list(tt_sh, tt_rows)
read_excel_rows_to_list(ca_sh, ca_rows)
tt_cells = remove_empty_rows(tt_rows)
ca_cells = remove_empty_rows(ca_rows)
# %%
# Create list of tuples. Each tuple corresponds to one row, skipping headers. 
tt_data = []
ca_data = []
for i in range(1, len(tt_cells)):
    tt_data.append(tt_cells[i])
for i in range(1, len(ca_cells)):
    ca_data.append(ca_cells[i])

# %%
if __name__ == "__main__":
    # Open Chrome
    driver = webdriver.Chrome('/Users/matte/Downloads/chromedriver')

    # Open the website
    driver.get("https://www.cilsfirst.com/")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/form/center/div/table/tbody/tr[3]/td/input")))

    # Find the username and password fields and enter the credentials
    driver.find_element_by_xpath('/html/body/form/center/div/table/tbody/tr[1]/td[2]/input').send_keys(username)
    driver.find_element_by_xpath('/html/body/form/center/div/table/tbody/tr[2]/td[2]/input').send_keys(password)
    driver.find_element_by_xpath('/html/body/form/center/div/table/tbody/tr[3]/td/input').click()
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="side_8"]/a')))


    # Access individual time tracker cells and assign to variables
    for i in range(len(tt_data)):
        # check for correct date format
        # check_time_format(i[1])
        # check_time_format(i[2])

        # assign variables
        date = str(tt_data[i][0])
        time_begun = str(tt_data[i][1])
        time_ended = str(tt_data[i][2])
        projects = tt_data[i][3].strip()
        funding_source = tt_data[i][4]
        note = tt_data[i][5]

        # enter first time tracker record
        if i == 0:
            # Find the time tracker button and click it
            i_path = driver.find_element_by_xpath('//*[@id="side_8"]/a').click()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="content_div"]/div[2]/table/tbody/tr[4]/td/a')))
            
            # Find 'add new time tracker record' button and click it
            add_new_path = driver.find_element_by_xpath('//*[@id="content_div"]/div[2]/table/tbody/tr[4]/td/a').click()

            # Enter date
            contact_date = driver.find_element_by_xpath('//*[@id="in_304clone"]').send_keys(date)

            # Enter time begun
            time_begun = driver.find_element_by_xpath('//*[@id="in_308"]').send_keys(time_begun)

            # Time ended HH:MM
            time_ended = driver.find_element_by_xpath('//*[@id="in_309"]').send_keys(time_ended)

            # Projects dropdown
            projects = driver.find_element_by_xpath('//*[@id="in_526"]').send_keys(projects)

            # Enter funding source dropdown
            funding_source = driver.find_element_by_xpath('//*[@id="in_313"]').send_keys(funding_source)

            # Enter social worker's note in text box
            staff_comments = driver.find_element_by_xpath('//*[@id="in_312"]').send_keys(note)

            # Click save record
            save_record = driver.find_element_by_xpath('//*[@id="sub"]').click()
        
        # enter remaining time tracker records
        else:
            # find add a new record (from 'record successfully saved') and click it
            add_new_path = driver.find_element_by_xpath('//*[@id="data_form"]/table[1]/tbody/tr[4]/td/input[3]').click()

            # Enter date
            contact_date = driver.find_element_by_xpath('//*[@id="in_304clone"]').send_keys(date) 

            # Enter time begun
            time_begun = driver.find_element_by_xpath('//*[@id="in_308"]').send_keys(time_begun)

            # Time ended HH:MM
            time_ended = driver.find_element_by_xpath('//*[@id="in_309"]').send_keys(time_ended)

            # Projects dropdown
            projects = driver.find_element_by_xpath('//*[@id="in_526"]').send_keys(projects)

            # Enter funding source dropdown
            funding_source = driver.find_element_by_xpath('//*[@id="in_313"]').send_keys(funding_source)

            # Enter social worker's note in text box
            staff_comments = driver.find_element_by_xpath('//*[@id="in_312"]').send_keys(note)

            # Click save record
            save_record = driver.find_element_by_xpath('//*[@id="sub"]').click()
    # Close the browser
    driver.quit()

    # Open new browser instance
    # Open Chrome
    driver = webdriver.Chrome('/Users/matte/Downloads/chromedriver')

    # Open the website
    driver.get("https://www.cilsfirst.com/")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/form/center/div/table/tbody/tr[3]/td/input")))

    # Find the username and password fields and enter the credentials
    driver.find_element_by_xpath('/html/body/form/center/div/table/tbody/tr[1]/td[2]/input').send_keys(username)
    driver.find_element_by_xpath('/html/body/form/center/div/table/tbody/tr[2]/td[2]/input').send_keys(password)
    driver.find_element_by_xpath('/html/body/form/center/div/table/tbody/tr[3]/td/input').click()
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="side_8"]/a')))

    for i in range(len(ca_data)):
        date = str(ca_data[i][0])
        hours = str(ca_data[i][1])
        time_begun = str(ca_data[i][2])
        time_ended = str(ca_data[i][3])
        issue_area = ca_data[i][4].strip()
        projects = ca_data[i][5].strip()
        service_program = ca_data[i][6].strip()
        priority_area = ca_data[i][7]
        funding_source = ca_data[i][8]
        note = ca_data[i][9]

        # Enter first community activity record
        if i == 0:
            # Find the Community Activities button and click it
            i_path = driver.find_element_by_xpath('//*[@id="side_4"]/a').click()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="content_div"]/div[2]/table/tbody/tr[4]/td/a')))
            
            # Find 'add new community activities record' button and click it
            add_new_path = driver.find_element_by_xpath('//*[@id="content_div"]/div[2]/table/tbody/tr[4]/td/a').click()

            # Enter date
            contact_date = driver.find_element_by_xpath('//*[@id="in_244clone"]').send_keys(date)

            # Enter time begun
            time_begun = driver.find_element_by_xpath('//*[@id="in_530"]').send_keys(time_begun)
            

            # Time ended HH:MM
            time_ended = driver.find_element_by_xpath('//*[@id="in_531"]').send_keys(time_ended)
            

            # Hours (float)
            hours = driver.find_element_by_xpath('//*[@id="in_250"]').send_keys(hours)
            

            # Issue area dropdown 
            issue_area = driver.find_element_by_xpath('//*[@id="in_247"]').send_keys(issue_area)
            

            # Projects dropdown
            projects = driver.find_element_by_xpath('//*[@id="in_258"]').send_keys(projects)
            

            # service program dropdown
            service_program = driver.find_element_by_xpath('//*[@id="in_248"]').send_keys(service_program)
            

            # priority area dropdown
            priority_area = driver.find_element_by_xpath('//*[@id="in_338"]').send_keys(priority_area)

            # Enter funding source dropdown
            funding_source = driver.find_element_by_xpath('//*[@id="in_303"]').send_keys(funding_source)
            

            # Enter social worker's note in text box
            staff_comments = driver.find_element_by_xpath('//*[@id="in_255"]').send_keys(note)
            

            # Click save record
            save_record = driver.find_element_by_xpath('//*[@id="sub"]').click()

        # Enter remaining community activities records
        else:
            # find add a new record (from 'record successfully saved') and click it
            add_new_path = driver.find_element_by_xpath('//*[@id="data_form"]/table[1]/tbody/tr[4]/td/input[3]').click()

            # Enter date
            contact_date = driver.find_element_by_xpath('//*[@id="in_244clone"]').send_keys(date)

            # Enter time begun
            time_begun = driver.find_element_by_xpath('//*[@id="in_530"]').send_keys(time_begun)
            
            # Time ended HH:MM
            time_ended = driver.find_element_by_xpath('//*[@id="in_531"]').send_keys(time_ended)
    
            # Hours (float)
            hours = driver.find_element_by_xpath('//*[@id="in_250"]').send_keys(hours)

            # Issue area dropdown 
            issue_area = driver.find_element_by_xpath('//*[@id="in_247"]').send_keys(issue_area)
            
            # Projects dropdown
            projects = driver.find_element_by_xpath('//*[@id="in_258"]').send_keys(projects)
            
            # service program dropdown
            service_program = driver.find_element_by_xpath('//*[@id="in_248"]').send_keys(service_program)
            
            # priority area dropdown
            priority_area = driver.find_element_by_xpath('//*[@id="in_338"]').send_keys(priority_area)

            # Enter funding source dropdown
            funding_source = driver.find_element_by_xpath('//*[@id="in_303"]').send_keys(funding_source)

            # Enter social worker's note in text box
            staff_comments = driver.find_element_by_xpath('//*[@id="in_255"]').send_keys(note)
            
            # Click save record
            save_record = driver.find_element_by_xpath('//*[@id="sub"]').click()
    
    # Close the driver
    driver.quit()