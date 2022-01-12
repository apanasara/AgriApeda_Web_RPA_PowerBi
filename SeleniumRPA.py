from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
import csv
import pandas as pd

driver = webdriver.Chrome(executable_path="D:\RPA\chromedriver_win32\chromedriver.exe")
ignored_exceptions=(NoSuchElementException,StaleElementReferenceException,)

driver.implicitly_wait(30)
driver.get("http://www.agriexchange.apeda.gov.in/IndExp/PortNew.aspx")
assert "India Export Statistics" in driver.title
driver.find_element_by_id("RadioButtonList3_1").click()
driver.find_element_by_id("USMill").click()
print('USMill selected')


#driver.find_element_by_xpath("//select[@id='ListBoxState']/option[1]").click()

select = Select(driver.find_element_by_id('ListBoxYear'))

for y in range(1,10):
    select.select_by_index(y)
    #driver.find_element_by_xpath("//select[@id='ListBoxYear']/option["+str(y)+"]").click()
    '''ActionChains(driver) \
        .key_down(Keys.CONTROL) \
        .click(element) \
        .key_up(Keys.CONTROL) \
        .perform()'''


for c in [0]:#range(230,244):
    print (" ")
    select = Select(driver.find_element_by_id('ListBoxCountry'))
    select.deselect_all()
    select.select_by_index(c)

    cntry= driver.find_element_by_xpath("//select[@id='ListBoxCountry']/option["+str(c+1)+"]").text    
    driver.find_element_by_css_selector("input[type='submit'][value='Submit']").click()
    print(f'submitted {c}...')
    
    try:
        while True:
            WebDriverWait(driver, 10).until(lambda driver: driver.execute_script('return document.readyState') == 'complete')
            print('document ready...')
            element_present = EC.presence_of_element_located((By.XPATH, "//td[contains(text(),'Total')]"))
            WebDriverWait(driver, 60,ignored_exceptions=ignored_exceptions).until(element_present)
            print(f'received {c}...')
            reportCountry = driver.find_element_by_xpath("//span[@id='countryregionname']")
            print(cntry)
            if cntry in reportCountry.text:
                print(reportCountry.text)
                table = driver.find_element_by_id("DataGrid1")
                filePath = 'D:\\Data Science\\Export Data India\\Source\\Country\\'+cntry+'.csv'
                htmlStr=str(table.get_attribute('outerHTML'))
                df=pd.read_html(htmlStr)[0]
                df.to_csv(filePath, index = True) 
                print('table exported to : '+filePath)
                break 

    except TimeoutException:
        print (f"Loading took too much time for {cntry}!")

driver.close()