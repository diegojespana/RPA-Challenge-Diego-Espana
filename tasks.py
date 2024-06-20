import os
import time
import requests
import shutil
from robocorp.tasks import task
from RPA.Excel.Files import Files
from robocorp.tasks import get_output_dir
from RPA.Browser.Selenium import Selenium
from selenium.webdriver.common.keys import Keys

@task
def minimal_task():
    #Define search phrase
    search_phrase = "NBA"

    #check if output excel file exists
    if os.path.exists("Results.xlsx"):
        os.remove("Results.xlsx")
    #Create output folder for downloaded images
    output_folder = get_output_dir()
    output_folder = os.path.join(output_folder, "Images")
    os.makedirs(output_folder, exist_ok=True)

    #Open website
    browser = Selenium()
    browser.open_available_browser('https://apnews.com/')
    #Search phrase
    browser.wait_until_element_is_visible("//button[@class='SearchOverlay-search-button']")
    browser.click_element_when_visible("//button[@class='SearchOverlay-search-button']")
    browser.wait_until_element_is_visible('name:q')
    browser.input_text('name:q', search_phrase +Keys.ENTER)

    #Filter by category
    browser.wait_until_element_is_visible("//div[@class='SearchResultsModule-results']", timeout=15)
    browser.click_element_when_visible("//div[@class='SearchFilter']")
    browser.click_element_when_visible('xpath: //*[contains(text(), "Stories")]')
    time.sleep(2)
    #Sort by newest
    browser.wait_until_element_is_visible("//div[@class='SearchResultsModule-filters-selected-heading']")
    browser.click_element_when_visible("//select[@class='Select-input']")
    browser.click_element_when_visible('xpath: //*[contains(text(), "Newest")]')
    time.sleep(5)
    #collect search results
    result_locator = "//div[@class='SearchResultsModule-results']/bsp-list-loadmore/div[2]/div"
    search_results = browser.find_elements(result_locator)
    print(search_results)

    results = {'Heading':[],'Link':[],'Description':[],'Date':[],'Image':[],'Count':[],'Money':[]}
    try:
        for result_index,elements in enumerate(search_results,1):
            #clean variables
            img = ''
            src = ''
            img_name = ''
            money = "False"
            times_in_title = 0
            times_in_description = 0

            if browser.is_element_visible(result_locator+f'[{result_index}]//bsp-custom-headline/div/a/span'):
                results['Heading'].append(browser.get_text(result_locator+f'[{result_index}]//bsp-custom-headline/div/a/span'))
                title = browser.get_text(result_locator+f'[{result_index}]//bsp-custom-headline/div/a/span')
                times_in_title = title.count(search_phrase)
                if "$" in title:
                    money = "True"
                elif "dollar" in title:
                    money = "True"
                elif "USD" in title:
                    money = "True"
            if browser.is_element_visible(result_locator+f'[{result_index}]/div/div/a'):
                results['Link'].append(browser.get_element_attribute(result_locator+f'[{result_index}]/div/div/a','href'))
            elif browser.is_element_visible(result_locator+f'[{result_index}]/div/div/div/a'):
                results['Link'].append(browser.get_element_attribute(result_locator+f'[{result_index}]/div/div/div/a','href'))
            if browser.is_element_visible(result_locator+f'[{result_index}]/div/div/div/a/span'):
                results['Description'].append(browser.get_text(result_locator+f'[{result_index}]/div/div/div/a/span'))
                description = browser.get_text(result_locator+f'[{result_index}]/div/div/div/a/span')
                times_in_description = description.count(search_phrase)
                if "$" in description:
                    money = "True"
                elif "dollar" in description:
                    money = "True"
                elif "USD" in description:
                    money = "True"
            if browser.is_element_visible(result_locator+f'[{result_index}]//div/div/bsp-timestamp/span/span'):
                results['Date'].append(browser.get_text(result_locator+f'[{result_index}]//div/div/bsp-timestamp/span/span'))
            if browser.is_element_visible(result_locator+f'[{result_index}]/div/div/a/picture/img'):
                img = browser.get_text(result_locator+f'[{result_index}]//bsp-custom-headline/div/a/span')
                src = browser.get_element_attribute(result_locator+f'[{result_index}]/div/div/a/picture/img','src')
                response = requests.get(src, stream=True)
                img_name = img + '.jpeg'
                results['Image'].append(img_name)
                output = os.path.join(output_folder, img_name)
                with open(output, 'wb') as out_file:
                    shutil.copyfileobj(response.raw, out_file)
                del response

            total_times = times_in_title + times_in_description
            results['Count'].append(total_times)
            results['Money'].append(money)


    except:
        print(Exception)
    finally:
        excel = Files()
        wb = excel.create_workbook()
        wb.create_worksheet('Results')
        excel.append_rows_to_worksheet(results,header=True,name='Results')
        wb.save(os.path.join(get_output_dir(),'Results.xlsx'))
    