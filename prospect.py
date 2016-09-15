#! /usr/bin/env python3

#Module is used to parse a list to prospect now with the intentions of producing a list 
#of owner addresses. This is the middleman list between the REA pull and the Lexis Nexis 
#parse. A data list may be imported directly here and not with an REA pull, however, 
#reimports to REA will be impossible with this current build if property keys are not 
#included in the list provided to this module. Driven by Selenium operations currently
#
#Required format for non REA used import is columns in the order of:
#
# address_1, city, state, zip_code

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException  
import bs4
import time


#Check if element is existant on a page
def assert_element_exists(xpath, browser):
    try:
        browser.find_element_by_xpath(xpath)
        return True
    except NoSuchElementException:
        return False

#Start of browser automation. This function handles all login and data point to each page
#that contains relevant data on which it will call the parse function which will append 
#owner information to the .xlsx file. 

def get_html(user, passwd, address_list):
    print('\nStarting ProspectNow crawl...')

    object_array = []
    options = webdriver.ChromeOptions()
    options.add_argument("--user-data-dir=/Users/bryan/Library/Application Support/Google/Chrome/Default/")
    options.add_argument("start-maximized")
    browser = webdriver.Chrome(chrome_options=options)

    #Root page for search 
    browser.get('https://www.prospectnow.com/search-commercial-real-estate/')

    #Login process
    username = browser.find_element_by_id("emailAddress2")
    password = browser.find_element_by_id("lpassword")

    #username.send_keys(user)
    #password.send_keys(passwd)
    browser.find_element_by_xpath("//*[@id=\"ExistingMemberLogin\"]/p/input").click()
    time.sleep(4)

    #Navigate to address input
    browser.find_element_by_xpath("//*[@id=\"step1_2\"]").click()
    browser.find_element_by_xpath("//*[@id=\"step1\"]/div[7]/a").click()
    time.sleep(3)

    #Input addresses
    text_area = browser.find_element_by_xpath("//*[@id=\"step3_addresses\"]")
    for address in address_list:
        text_area.send_keys(address+'\n')

    browser.find_element_by_xpath("//*[@id=\"step2\"]/a[2]").click()
    time.sleep(30)

    #Move through pages and return property dictionary list from each page
    print('Saving property data...')
    last_page = False
    #array = browser.execute_script('return global_results_array;')
    #object_array.append(array[1])
    time.sleep(3)

    while last_page is False:
        
        time.sleep(3)
        browser.execute_script('nextPage();')
        if not (assert_element_exists("//*[@id=\"pag\"]/li[6]/a", browser) or assert_element_exists("//*[@id=\"pag\"]/li[7]/a", browser)):
            last_page = True
            array = browser.execute_script('return global_results_array;')
            print('Finished PropspectNow crawl.')
            browser.quit()
            return array







