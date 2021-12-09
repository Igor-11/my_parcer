from RPA.Browser.Selenium import Selenium
from constants.urls import HOME_PAGE
from constants.locators import DIVE_IN_BUTTON, AMOUNT_INFO, AGENCY_NAME_INFO, AGENCY_LINK, UII_COLUMN, SECOND_COLUMN, THIRD_COLUMN, FOURTH_COLUMN, FIFTH_COLUMN, SIXTH_COLUMN, SEVENTH_COLUMN
import time
from RPA.Excel.Files import Files
from config import AGENCY_NAME


browser = Selenium()
def open_browser(url):
    browser.open_available_browser(url)

def click_button(button_locator):
    browser.click_element_if_visible(button_locator)

def main():
    open_browser(HOME_PAGE)
    click_button(DIVE_IN_BUTTON)
    save_agency_info()
    select_agency(AGENCY_NAME)
    save_individual_investments()

def save_agency_info():
    time.sleep(5)
    amounts = [element.get_attribute('innerHTML') for element in browser.find_elements(AMOUNT_INFO)]
    agencies = [element.get_attribute('innerHTML') for element in browser.find_elements(AGENCY_NAME_INFO)]
    file = Files()
    workbook = file.create_workbook('./output/Agencies.xlsx')
    workbook.create_worksheet('Agencies')
    for index in range(len(amounts)):
         workbook.set_cell_value(index+1,1,agencies[index])
         workbook.set_cell_value(index+1,2,amounts[index])
    file.save_workbook()

def select_agency(agency_name):
    time.sleep(3)
    link=browser.find_element(AGENCY_LINK(agency_name))
    browser.click_element_if_visible(link)

def save_individual_investments():
    time.sleep(10)
    uii = [element.get_attribute('innerHTML') for element in browser.find_elements(UII_COLUMN)]
    second = [element.get_attribute('innerHTML') for element in browser.find_elements(SECOND_COLUMN)]
    third = [element.get_attribute('innerHTML') for element in browser.find_elements(THIRD_COLUMN)]
    fourth = [element.get_attribute('innerHTML') for element in browser.find_elements(FOURTH_COLUMN)]
    fifth = [element.get_attribute('innerHTML') for element in browser.find_elements(FIFTH_COLUMN)]
    sixth = [element.get_attribute('innerHTML') for element in browser.find_elements(SIXTH_COLUMN)]
    seventh = [element.get_attribute('innerHTML') for element in browser.find_elements(SEVENTH_COLUMN)]
    file = Files()
    file.open_workbook('./output/Agencies.xlsx')
    file.create_worksheet('Individual investments', exist_ok=True)
    for index in range(len(uii)):
        file.workbook.set_cell_value(index+1,1,uii[index])
        file.workbook.set_cell_value(index+1,2,second[index])
        file.workbook.set_cell_value(index+1,3,third[index])
        file.workbook.set_cell_value(index+1,4,fourth[index])
        file.workbook.set_cell_value(index+1,5,fifth[index])
        file.workbook.set_cell_value(index+1,6,sixth[index])
        file.workbook.set_cell_value(index+1,7,seventh[index])
    file.save_workbook()

    
    
if __name__ == '__main__':
    main()
