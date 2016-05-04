#! /usr/bin/env python3

#Module is used to parse the Lexis network for emails and phone numbers of the returned
#ProspectNow contacts

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException  
import bs4
import time
import codecs
import re
import os

#Webdriver
driver = ''

#Path to downloads
path_to_folder = '/users/bryan/downloads/'

#Check if element is existant on a page
def assert_element_exists(xpath, browser):
    try:
        browser.find_element_by_xpath(xpath)
        return True
    except NoSuchElementException:
        return False

#LOGIN to Lexis Root
def login(first, last, email, location):
    print('\nLogging into LexisNexis...')
    global driver
    driver = webdriver.Chrome()
    driver.get('https://w3.nexis.com/clients/marcusmillichapregistration/')
    time.sleep(1)
    first = driver.find_element_by_xpath('//*[@id=\"firstName\"]')
    last  = driver.find_element_by_xpath('//*[@id=\"lastName\"]')
    email  = driver.find_element_by_xpath('//*[@id=\"open_one\"]')
    location = Select(driver.find_element_by_xpath('//*[@id=\"location\"]'))
    location.select_by_visible_text('Newport Beach')
    first.send_keys('Bryan')
    last.send_keys('Wheeler')
    email.send_keys('bryan.wheeler@marcusmillichap.com')
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id=\"content\"]/div[2]/div/form/ol/li[6]/input[1]').click()
    driver.switch_to_frame('mainFrame')
    driver.switch_to_frame('fr_feature')

#Send address fields to Lexis form and parse to return phone and email
def send_data(address, city, state, zip_code):
    try:
        driver.switch_to_frame('mainFrame')
        driver.switch_to_frame('fr_feature')
    
    except:
        print(address, city, state, zip_code)

        address_field = driver.find_element_by_xpath("//*[@id=\"MainContent_Address1\"]")
        city_field = driver.find_element_by_xpath("//*[@id=\"MainContent_City\"]")
        state_sel = Select(driver.find_element_by_xpath('//*[@id=\"MainContent_State_stateList\"]'))
        zip_code_field = driver.find_element_by_xpath("//*[@id=\"MainContent_Zip5\"]")

        address_field.send_keys(address)
        city_field.send_keys(city)
        state_sel.select_by_visible_text(state)
        zip_code_field.send_keys(zip_code)
        driver.find_element_by_xpath('//*[@id=\"MainContent_formSubmit_searchButton\"]').click()


#Access first Lexis name link
def access_owner():
    try:
        #Click on owner
        if assert_element_exists('//*[@id=\"spanNames1_0\"]', driver):
            elem = driver.find_element_by_xpath('//*[@id=\"spanNames1_0\"]')
            action = webdriver.ActionChains(driver)
            action.move_to_element_with_offset(elem, 5, 5)
            action.click()
            action.perform()
            time.sleep(2)
        else:
            elem = driver.find_element_by_xpath('//*[@id=\"spanNames2_0\"]')
            action = webdriver.ActionChains(driver)
            action.move_to_element_with_offset(elem, 5, 5)
            action.click()
            action.perform()
            time.sleep(2)

        #Download the data file
        driver.find_element_by_id('MainContent_toolbar_deliveryDisk').click()
        time.sleep(2)

        #Switch to download window and initiate download then return driver to root page and clear fields
        for handle in driver.window_handles:
            driver.switch_to_window(handle)

        driver.find_element_by_xpath('//*[@id=\"MainContent_ButtonDownloadPreview\"]').click()
        time.sleep(10)
        driver.find_element_by_xpath('//*[@id=\"MainContent_ButtonContinue\"]').click()
        time.sleep(2)
        driver.close()
        for handle in driver.window_handles:
            driver.switch_to_window(handle)
        driver.back()
        driver.back()

        #Clear search fields
        driver.switch_to_frame('mainFrame')
        driver.switch_to_frame('fr_feature')
        address = driver.find_element_by_xpath('//*[@id=\"MainContent_Address1\"]').clear()
        city = driver.find_element_by_xpath('//*[@id=\"MainContent_City\"]').clear()
        zip_code = driver.find_element_by_xpath("//*[@id=\"MainContent_Zip5\"]").clear()
        time.sleep(1)

        #PARSE FILE HERE

    #If there is no person info, switch back to root page
    except:
        driver.back()
        driver.switch_to_frame('mainFrame')
        driver.switch_to_frame('fr_feature')
        address = driver.find_element_by_xpath('//*[@id=\"MainContent_Address1\"]').clear()
        city = driver.find_element_by_xpath('//*[@id=\"MainContent_City\"]').clear()
        zip_code = driver.find_element_by_xpath("//*[@id=\"MainContent_Zip5\"]").clear()
        time.sleep(2)
        pass

        #PARSE FILE HERE

#Parse for phone numbers and return string
def __parse_phones(parser):
    node_1 = parser.find('span', text='Phone')
    node_2 = node_1.find_next('span')
    node_3 = node_2.find_next('span')
    node_4 = node_3.find_next('span')
    phones = node_4.find_next('span')
    return str(phones).replace('<span class=\"c11\">', '').replace('<br>', ', ').replace('</br>', '').replace('</span>', '')

#Parse for emails and return string
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
        return 'No Emails Found'

#Parse downloaded file and delete
def parse_file():
    contact = {}
    for fn in os.listdir(path_to_folder):
        if fn.endswith('.html'):
            parser = bs4.BeautifulSoup(codecs.open(path_to_folder + fn, 'r').read(), 'html.parser')
            phones = __parse_phones(parser)
            emails = __parse_emails(parser)
            contact.update({'emails':emails, 'phones':phones})
            os.remove(path_to_folder + fn)
    return contact

#Search business directly from Lexis Nexis. This is a full funtcion that handles both the download
#and parse for each Property object. This search will initiate if prop.owner_is_businees is true
def business_search(prop):
    #set the biz name
    business = prop.owner_first
    try:
        #run the search to first page
        driver.find_element_by_xpath("//*[@id=\"menuLevel0Item1\"]/a").click()
        time.sleep(1)
        driver.find_element_by_xpath("//*[@id=\"menuLevel3Item0\"]/a").click()
        time.sleep(1)
        biz_in = driver.find_element_by_xpath("//*[@id=\"MainContent_CompanyName\"]")
        biz_in.send_keys(business)
        driver.find_element_by_xpath("//*[@id=\"MainContent_formSubmit_searchButton\"]").click()
        time.sleep(3)

        #click first link
        driver.find_element_by_xpath("//*[@id=\"resultscontent\"]/tbody/tr[1]/td[2]/a").click()

        if assert_element_exists('//*[@id=\"Executives\"]/div/table/tbody/tr[1]/td[4]/a', driver):
            driver.find_element_by_xpath("//*[@id=\"Executives\"]/div/table/tbody/tr[1]/td[4]/a").click()
            time.sleep(2)

            #Download the data file
            driver.find_element_by_id('MainContent_toolbar_deliveryDisk').click()
            time.sleep(2)

            #Switch to download window and initiate download then return driver to root page and clear fields
            for handle in driver.window_handles:
                driver.switch_to_window(handle)

                driver.find_element_by_xpath('//*[@id=\"MainContent_ButtonDownloadPreview\"]').click()
                
        else:
            pass

    except:
        time.sleep(1)
        





