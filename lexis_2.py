#! /usr/bin/env python3

#Module is used to parse the Lexis network for emails and phone numbers of the returned
#ProspectNow contacts

from selenium import webdriver
#from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException  
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import bs4
import time
import codecs
import re
import os

#Webdriver global amd property global
driver = ''

#Path to downloads folder
path_to_folder = '/users/bryan/downloads/'


#Check if element on a page exists by xpath
def assert_element_exists(xpath, browser):
    try:
        browser.find_element_by_xpath(xpath)
        return True
    except NoSuchElementException:
        return False

#Reset the driver to the root page
def reset():
    driver.get('https://w3.nexis.com/clients/marcusmillichap/')

#Quit the browser
def quit():
    driver.quit()

#Login into LexisNexis using the arguments as credentials and navigate to the search page
def login(first_, last_, email_, location_):
    print('\nLogging into LexisNexis...')
    
    #Set webdriver global
    global driver
    driver = webdriver.Chrome()

    #Go to login page
    driver.get('https://w3.nexis.com/clients/marcusmillichapregistration/')
    time.sleep(1)

    #Set input fields
    first = driver.find_element_by_xpath('//*[@id=\"firstName\"]')
    last  = driver.find_element_by_xpath('//*[@id=\"lastName\"]')
    email  = driver.find_element_by_xpath('//*[@id=\"open_one\"]')
    location = Select(driver.find_element_by_xpath('//*[@id=\"location\"]'))
    
    #Send data to input fields 
    location.select_by_visible_text(location_)
    first.send_keys(first_)
    last.send_keys(last_)
    email.send_keys(email_)
    time.sleep(1)

    #Click submit
    driver.find_element_by_xpath('//*[@id=\"content\"]/div[2]/div/form/ol/li[6]/input[1]').click()

#Clear fields on form (arguement 'b' for business form and 'p' for person form)
def clear_form(form_type):
    try:
        if form_type == 'b':
            driver.find_element_by_xpath("//*[@id=\"MainContent_CompanyName\"]").clear()
            driver.find_element_by_xpath("//*[@id=\"MainContent_Address1\"]").clear()
            driver.find_element_by_xpath("//*[@id=\"MainContent_City\"]").clear()
            driver.find_element_by_xpath("//*[@id=\"MainContent_Zip5\"]").clear()
        elif form_type == 'p':
            driver.find_element_by_xpath("//*[@id=\"MainContent_LastName\"]").clear()
            driver.find_element_by_xpath("//*[@id=\"MainContent_FirstName\"]").clear()
            driver.find_element_by_xpath("//*[@id=\"MainContent_Address1\"]").clear()
            driver.find_element_by_xpath("//*[@id=\"MainContent_City\"]").clear()
            driver.find_element_by_xpath("//*[@id=\"MainContent_Zip5\"]").clear()
    except:
        pass

#If ownership is an individual, run individual crawl branch 1. If not, run business crawl branch 1
def run_crawl(prop):


    print('Running scrape on ' + prop.owner_first + ' ' + prop.owner_last)

    #Switch to content frame
    driver.switch_to_frame('mainFrame')
    driver.switch_to_frame('fr_feature')

    #Determine ownership for business
    if prop.owner_is_business is True:
        
        #Navigate to biz report area
        biz_tab = driver.find_element_by_xpath("//*[@id=\"menuLevel0Item1\"]/a").click()
        time.sleep(.5)
        report_link = driver.find_element_by_xpath("//*[@id=\"menuLevel3Item0\"]/a").click()
        time.sleep(.5)

        #Clear all input fields
        clear_form('b')

        #Set input vars
        company_field = driver.find_element_by_xpath("//*[@id=\"MainContent_CompanyName\"]")
        address_field = driver.find_element_by_xpath("//*[@id=\"MainContent_Address1\"]")
        city_field = driver.find_element_by_xpath("//*[@id=\"MainContent_City\"]")
        state_sel = Select(driver.find_element_by_xpath('//*[@id=\"MainContent_State_stateList\"]'))
        zip_code_field = driver.find_element_by_xpath("//*[@id=\"MainContent_Zip5\"]")

        #Send individual address info and click submit
        company_field.send_keys(prop.owner_first)
        address_field.send_keys(prop.owner_address)
        city_field.send_keys(prop.owner_city)
        state_sel.select_by_visible_text(prop.get_state(False))
        zip_code_field.send_keys(prop.owner_zip)
        driver.find_element_by_xpath("//*[@id=\"MainContent_formSubmit_searchButton\"]").click()
        time.sleep(.5)
        
        #Navigate to company executive person report if exists
        __move_to_b_report()
        time.sleep(.5)
        __move_to_b_person()
        time.sleep(.5)

        #Return parsed contact info
        results = __run_download()

        if results:
            return results
        else:
            return {'phones':'', 'emails':''}


    #Navigate individual search
    else:

        #Clear all input fields
        clear_form('p')

        #Set input field variables
        last_name = driver.find_element_by_xpath("//*[@id=\"MainContent_LastName\"]")
        first_name = driver.find_element_by_xpath("//*[@id=\"MainContent_FirstName\"]")
        address_field = driver.find_element_by_xpath("//*[@id=\"MainContent_Address1\"]")
        city_field = driver.find_element_by_xpath("//*[@id=\"MainContent_City\"]")
        state_sel = Select(driver.find_element_by_xpath('//*[@id=\"MainContent_State_stateList\"]'))
        zip_code_field = driver.find_element_by_xpath("//*[@id=\"MainContent_Zip5\"]")

        #Send individual address info and click submit
        last_name.send_keys(prop.owner_last)
        first_name.send_keys(prop.owner_first)
        address_field.send_keys(prop.owner_address)
        city_field.send_keys(prop.owner_city)
        state_sel.select_by_visible_text(prop.get_state(False))
        zip_code_field.send_keys(prop.owner_zip)
        driver.find_element_by_xpath('//*[@id=\"MainContent_formSubmit_searchButton\"]').click()
        time.sleep(.5)

        #Navigate to individual report 1 or 2 if not 1
        __move_to_p_report()
        time.sleep(.5)

        #Return parsed contact info
        results = __run_download()

        if results:
            return results
        else:
            return {'phones':'', 'emails':''}





#############BUSINESS BRANCH FUNCTIONS################################



#Move to the business report if availiable. If not, return to search page 
def __move_to_b_report():

    if assert_element_exists("//*[@id=\"resultscontent\"]/tbody/tr[1]/td[2]/a", driver):
        driver.find_element_by_xpath("//*[@id=\"resultscontent\"]/tbody/tr[1]/td[2]/a").click()

    else:
        driver.back()



#Move to person report if company executive exists, if not move back to main search
def __move_to_b_person():

    if assert_element_exists("//*[@id=\"Executives\"]/div/table/tbody/tr[1]/td[4]/a", driver):
        driver.find_element_by_xpath("//*[@id=\"Executives\"]/div/table/tbody/tr[1]/td[4]/a").click()

    else:
        driver.back()
        driver.back()



#############INDIVIDUAL BRANCH FUNCTIONS################################



#Move to first person selection if it exists, if not move to second selection, if not move back
def __move_to_p_report():

    #Click on owner 1 if available
    if assert_element_exists('//*[@id=\"spanNames1_0\"]', driver):
        elem = driver.find_element_by_xpath('//*[@id=\"spanNames1_0\"]')
        action = webdriver.ActionChains(driver)
        action.move_to_element_with_offset(elem, 5, 5)
        action.click()
        action.perform()

    #Click on owner 2 if available
    elif assert_element_exists('//*[@id=\"spanNames2_0\"]', driver):
        elem = driver.find_element_by_xpath('//*[@id=\"spanNames2_0\"]')
        action = webdriver.ActionChains(driver)
        action.move_to_element_with_offset(elem, 5, 5)
        action.click()
        action.perform()

    else:
        driver.back()


#################DOWNLOAD HANDLER##############################################



#Run the download of the report HTML file and close the popup after download, returns the 
#owner information and reset the browser to the initial search page.
def __run_download():
    try:
        if assert_element_exists('//*[@id=\"DisplaySubjectSummary\"]', driver):
            print('Handeling download...')

            #Click download preview
            driver.find_element_by_xpath("//*[@id=\"MainContent_toolbar_deliveryDisk\"]").click()
            time.sleep(2)

            #Switch to resulting popup window 
            for handle in driver.window_handles:
                driver.switch_to_window(handle)

            #Click and initiate download
            driver.find_element_by_xpath('//*[@id=\"MainContent_ButtonDownloadPreview\"]').click()
            time.sleep(10)
            driver.find_element_by_xpath('//*[@id=\"MainContent_ButtonContinue\"]').click()
            time.sleep(2)

            #Parse the downloaded file, delete, and return the owner information
            contact = __parse_file()

            #Close the popup and reset window focus
            driver.close()

            #Switch back to main window
            for handle in driver.window_handles:
                driver.switch_to_window(handle)

            driver.get('https://w3.nexis.com/clients/marcusmillichap/')

            return contact

        else:

            driver.get('https://w3.nexis.com/clients/marcusmillichap/')

    except:

        driver.get('https://w3.nexis.com/clients/marcusmillichap/')




####################PARSE FUNCTIONS############################################



#Parse for phone numbers and return string. Function is only called by __parse_file()
def __parse_phones(parser):
    node_1 = parser.find('span', text='Phone')
    node_2 = node_1.find_next('span')
    node_3 = node_2.find_next('span')
    node_4 = node_3.find_next('span')
    phones = node_4.find_next('span')
    string = str(phones).replace('<span class=\"c11\">', '').replace('<br>', ', ').replace('</br>', '').replace('</span>', '')
    if string:
        return str(phones).replace('<span class=\"c11\">', '').replace('<br>', ', ').replace('</br>', '').replace('</span>', '')
    else:
        return ''


#Parse for emails and return string. Function is only called by __parse_file()
def __parse_emails(parser):
    emails = parser.find_all('span', text=re.compile('@'))
    if emails:
        ls = []
        if len(emails) == 1:
            for email in emails:
                ls.append(email.text)
            return ''.join(ls).lower()
        else:
            for email in emails:
                ls.append(email.text + ', ')
                string = ''.join(ls).lower()
                last_c = len(string) - 2

            return string[0:last_c]
    else:
        return ''

#Parse for name and return the string Function is only called by __pars_file()
def __parse_name(parser):
    node_1 = parser.find('span', text='Phone')
    node_2 = node_1.find_next('span').text
    return node_2


#Parse downloaded file and delete. Returns a dictionary of 'emails' and 'phones'. Only called by __run_download()
def __parse_file():
    contact = {}
    for fn in os.listdir(path_to_folder):
        if fn.endswith('.html'):
            parser = bs4.BeautifulSoup(codecs.open(path_to_folder + fn, 'r').read(), 'html.parser')
            phones = __parse_phones(parser)
            emails = __parse_emails(parser)
            name = __parse_name(parser)
            contact.update({'emails':emails, 'phones':phones, 'name':name})
            os.remove(path_to_folder + fn)
    if contact:
        return contact
    else:
        return {'emails':'problem here', 'phone':'problem here', 'name':'problem here'}




































