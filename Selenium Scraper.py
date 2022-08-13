import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException

driver = webdriver.Chrome()#initiating the driver
driver.set_window_size(800, 800)#adjasting the size of Chrome window
driver.get("http://www.aplusa-online.com/vis/v1/en/directory/q?oid=19008&lang=2")#get start page
letters_links = [link.get_attribute('href') for link in driver.find_elements_by_xpath('//*[@id="site-wrapper"]/div[5]/div/div[1]/a')]#create list of letters links

report = [] #emty list for companies info
for letter in letters_links: #iteration through letters
    driver.get(letter) # clicking on particular letter
    companies_links = [link.get_attribute('href') for link in driver.find_elements_by_class_name('flush')]#create list of links of companies
    for company in companies_links:#iteration through list of companies
        if type(company) == str: #checking of the company is a string because there are sometimes is None
            driver.get(company)
        else:
            continue
            #I am using try/except because some data is missing and to catch this exception and to add to report as None(you cannot do it with if/else)
        try:
            name = driver.find_element_by_xpath('//*[@id="vis__profile"]/div[2]/div/div/h1').text #get the name of company
            report.append(name)# append that name to report
        except NoSuchElementException:
            report.append('none')#if there are no info add none to the report
        try:
            city = driver.find_element_by_xpath('//*[@id="vis__profile"]/div[2]/div/div/div[2]/div[1]/div/span[3]').text
            report.append(city)
        except NoSuchElementException:
            report.append('none')
        try:
            country = driver.find_element_by_xpath(
                '//*[@id="vis__profile"]/div[2]/div/div/div[2]/div[1]/div/span[4]').text
            report.append(country)
        except NoSuchElementException:
            report.append('none')
        try:
            phone = driver.find_element_by_xpath('//*[@id="vis__profile"]/div[2]/div/div/div[2]/div[1]/p/span[1]').text
            report.append(phone)
        except NoSuchElementException:
            report.append('none')
        try:
            mail = driver.find_element_by_xpath('//*[@id="vis__profile"]/div[2]/div/div/div[2]/div[1]/p/a[1]').text
            report.append(mail)
        except NoSuchElementException:
            report.append('none')
        try:
            web = driver.find_element_by_xpath('//*[@id="vis__profile"]/div[2]/div/div/div[2]/div[1]/p/a[2]').text
            report.append(web)
        except NoSuchElementException:
            report.append('none')

result = [report[i:i + 6] for i in range(0, len(report), 6)]#make list of lists from the report
writer = pd.ExcelWriter('report.xlsx', engine='xlsxwriter')#working with pandas, creating writer
df = pd.DataFrame(result, index=range(6), columns=['name', 'city', 'country', 'phone', 'mail', 'web'])# framing data
df.to_excel(writer, sheet_name='aplus', index=False)#convert data frame to excel

driver.quit()#quit the driver
