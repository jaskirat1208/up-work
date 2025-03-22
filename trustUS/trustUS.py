import threading
import tkinter as tk
# from multiprocessing import freeze_support
from tkinter import filedialog as fd
from tkinter import ttk
import openpyxl
import requests
import time
import sv_ttk
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re
# Replace the required_category to your required category
required_category = 'shopping_fashion'

# Open the Excel file
workbook = openpyxl.load_workbook(
    'company.xlsx')
sheet = workbook.active

def scrape_info(soup, start_row):
    global info_text
    # workbook.save(f'Company-{required_category}_update.xlsx')

    business_links = soup.find_all('a', {'name': 'business-unit-card'})

    # Extract the href values
    hrefs = [link['href'] for link in business_links if 'href' in link.attrs]

    # Process each href
    for link in business_links:
        sheet[f'A{start_row}'].value = start_row-1
        company_name=link.find('p', {'class': re.compile('(^|.*)Typography.*')}).text.strip()
        sheet[f'B{start_row}'].value = company_name
        try:
            company_score=link.find('p', {"class": re.compile(".*styles_rating.*")}).text.strip()
            sheet[f'D{start_row}'].value = company_score
        except:
            sheet[f'D{start_row}'].value = "TrustScore 0|0 reviews"

        try:
            company_location = link.find('span', {"class": re.compile("(^|.*)styles_location.*")}).text.strip()
            sheet[f'E{start_row}'].value = company_location
        except:
            sheet[f'E{start_row}'].value = "Australia"

        href=link['href']
        company_domain=href.replace("/review/", "")
        sheet[f'C{start_row}'].value = company_domain
        business_url = f"https://au.trustpilot.com{href}"
        print(f"Processing {business_url}...")
        time.sleep(2)
        # Send a GET request to the individual business page
        response=requests.request("GET", business_url)
        response_sub =response.text

        # Check if the request was successful
        if response.status_code == 200:
            # Further processing can be done here with response_sub.content
            print(f"Successfully retrieved data from {business_url}")
            soup_sub = BeautifulSoup(response_sub, "html.parser")
            try:
                company_activity_card = soup_sub.find('span', {"class": re.compile("(^|.*)styles_companyActivity(.*|$)")})
                list_items = company_activity_card.find_all('span', {"class": re.compile("(^|.*)styles_listItem(.*|$)")})
                company_activiti_string= "".join([list.get_text().strip() + " / " for list in list_items])
            except:
                company_activiti_string= ""
            sheet[f'F{start_row}'].value = company_activiti_string
            start_row += 1

        else:
            print(f"Failed to retrieve data from {business_url}")


        print("\n")

def wait_and_get_element(driver, by, val):
    WebDriverWait(driver, 10000).until(
        EC.element_to_be_clickable((by, val))
    )
    elt = driver.find_element(by, val)
    return elt

def process_request():
    driver = webdriver.Chrome()  # You may need to adjust the path to your Chrome driver
    # Launch Chrome browser
    url = f"https://au.trustpilot.com/categories/{required_category}"
    driver.get(url)

    agree_buttons = wait_and_get_element(driver, By.CLASS_NAME, "onetrust-close-btn-handler")
    agree_buttons.click()
    try:
        # Try to find the element with data-pagination-button-last-link='true'
        page = wait_and_get_element(driver, By.XPATH, "//a[@data-pagination-button-last-link='true']").text
    except:
        try:
            page = wait_and_get_element(driver, By.XPATH, "//a[@data-pagination-button-4-link='true']").text
        except:
            try:
                page = wait_and_get_element(driver, By.XPATH, "//a[@data-pagination-button-3-link='true']").text
            except:
                try:
                    page = wait_and_get_element(driver, By.XPATH, "//a[@data-pagination-button-2-link='true']").text
                except:
                    page = wait_and_get_element(driver, By.XPATH, "//a[@data-pagination-button-1-link='true']").text

    page = int(page)
    # Find all pagination buttons
    for idx in range(0, page):
        print(f"Processing page {idx+1}")
        page_source = driver.page_source
        soup = BeautifulSoup(page_source, "html.parser")
        scrape_info(soup, start_row=idx * 20 + 2)
        # Check if it's not the last page
        if idx < page:
            button = wait_and_get_element(driver, By.XPATH,"//a[@data-pagination-button-next-link='true']")
            driver.execute_script("arguments[0].click();", button)
            # Wait for the new content to load
            WebDriverWait(driver, 20)

            # Optionally, add a short delay to ensure the content is fully loaded
            time.sleep(5)
        workbook.save(f'Company-{required_category}-{idx}.xlsx')
    driver.quit()  # Close the browser after scraping


def main():


    process_request()
    # Save the modified workbook

    try:
        workbook.save(f'Company-{required_category}.xlsx')
        # Close the workbook
        workbook.close()
    except FileNotFoundError:
        print("file cannot found!")
        return


if __name__ == '__main__':
    main()

