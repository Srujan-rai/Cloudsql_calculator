from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.common.exceptions import JavascriptException
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
import time
import os
import pandas as pd
import re
import csv
import glob
import json
import smtplib
from email.message import EmailMessage
import requests
from flask import Flask, request,jsonify
import pyperclip





def home_page(driver,actions):
        """Navigates to the pricing section."""
        driver.implicitly_wait(5)
        add_to_estimate_button = driver.find_element(By.XPATH, "//span[text()='Add to estimate']")
        add_to_estimate_button.click()
        time.sleep(5)
        div_element = driver.find_element(By.XPATH, "//div[@class='d5NbRd-EScbFb-JIbuQc PtwYlf' and @data-service-form='7']")
        actions.move_to_element(div_element).click().perform()
        time.sleep(2)
        print("✅ home page done")


def service_type(driver,actions,service_type):
    time.sleep(0.6)
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.6)
    
    pyautogui.typewrite('Service type')
    time.sleep(0.6)
    
    pyautogui.press('esc')
    time.sleep(0.6)
    
    actions.send_keys(Keys.ENTER).perform()
    time.sleep(0.6)
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.6)
    
    pyautogui.typewrite(service_type)
    time.sleep(0.6)
    
    pyautogui.press('esc')
    time.sleep(0.6)   
    actions.send_keys(Keys.ENTER).perform()
   
    
    print("✅ service type  selected")




def select_region(driver, actions, region):
    time.sleep(0.6)
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.6)
    
    pyautogui.typewrite('Region')
    time.sleep(0.6)
    
    pyautogui.press('esc')
    time.sleep(0.6)
    
    actions.send_keys(Keys.ENTER).perform()
    time.sleep(0.6)
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.6)
    
    pyautogui.typewrite(region)
    time.sleep(0.6)
    
    pyautogui.press('esc')
    time.sleep(0.6)   
    actions.send_keys(Keys.ENTER).perform()
   
    
    print("✅ Region selected")


def advanced_toggle_on(driver,actions):
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.6)
    
    pyautogui.typewrite('Advanced settings')
    time.sleep(0.6)
    
    pyautogui.press('esc')
    time.sleep(0.6)
    
    actions.key_down(Keys.SHIFT).send_keys(Keys.TAB).key_up(Keys.SHIFT).perform()
    time.sleep(0.6)
    actions.send_keys(Keys.ENTER).perform()
    print("✅ Advanced settings turned on")




def cloud_sql_edition(driver,actions,edition):
    time.sleep(0.6)
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.6)
    
    pyautogui.typewrite('Cloud SQL Edition')
    time.sleep(0.6)
    
    pyautogui.press('esc')
    time.sleep(0.6)
    
    actions.send_keys(Keys.ENTER).perform()
    time.sleep(0.6)
    
    for _ in range(2):
        actions.send_keys(Keys.TAB).perform()
        time.sleep(0.2)
    
    
    if edition=="Enterprise":
        pass
    
    elif edition=="Enterprise Plus":
        actions.send_keys(Keys.ARROW_RIGHT).perform()
    
    actions.send_keys(Keys.ENTER).perform()
   
    
    print("✅ Edition selected")


def Specify_usage_time(driver,actions):
    time.sleep(0.6)
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.6)
    
    pyautogui.typewrite('Specify usage time for each instance')
    time.sleep(0.6)
    
    pyautogui.press('esc')
    time.sleep(0.6)
    
    actions.send_keys(Keys.ENTER).perform()
    time.sleep(0.6)
    
    actions.key_down(Keys.SHIFT).send_keys(Keys.TAB).key_up(Keys.SHIFT).perform()

    
    actions.send_keys(Keys.ENTER).perform()
   
    
    print("✅ Usage time selected")


def instance(driver,actions,instance):
    time.sleep(0.6)
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.6)
    
    pyautogui.typewrite('Number of instances')
    time.sleep(0.6)
    
    pyautogui.press('esc')
    time.sleep(0.6)
    
    actions.send_keys(Keys.ENTER).perform()
    
    for _ in range(3):
        actions.send_keys(Keys.TAB).perform()
        time.sleep(0.2)
    
    actions.send_keys(instance).perform()
    time.sleep(0.3)
    
    for _ in range(2):
        actions.send_keys(Keys.TAB).perform()
        time.sleep(0.2)
    print("✅ Instance selected")




def usage_time(driver,actions,usage_time):
    time.sleep(0.6)
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.6)
    
    pyautogui.typewrite('Total instance usage time')
    time.sleep(0.6)
    
    pyautogui.press('esc')
    time.sleep(0.6)
    
    actions.send_keys(Keys.ENTER).perform()
    
    for _ in range(3):
        actions.send_keys(Keys.TAB).perform()
        time.sleep(0.2)
    time.sleep(0.2)
    actions.send_keys(usage_time).perform()
    time.sleep(0.2)
    for _ in range(2):
        actions.send_keys(Keys.TAB).perform()
        time.sleep(0.2)
    
    print("✅ Usage time selected")





def select_sql_instance_type(driver,actions,instance_type):
    time.sleep(0.6)
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.6)
    
    pyautogui.typewrite('Select SQL instance type')
    time.sleep(0.6)
    
    pyautogui.press('esc')
    time.sleep(0.6)
    
    actions.send_keys(Keys.ENTER).perform()
    for _ in range(2):
        actions.send_keys(Keys.TAB).perform()
        time.sleep(0.2)
    actions.send_keys(Keys.ENTER).perform()
    time.sleep(0.6)
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.6)
    instance_type=str(instance_type)
    pyautogui.typewrite(instance_type)
    time.sleep(0.6)
    
    pyautogui.press('esc')
    time.sleep(0.6)   
    actions.send_keys(Keys.ENTER).perform()
    actions.send_keys(Keys.ENTER).perform()
    time.sleep(0.3)
    print("✅ SQL instance type selected")




def vcpu_handle(driver,actions,vcpu,ram):
    if vcpu==0:
        print("Skipped Vcpu as it is default")
        pass
    else:
        time.sleep(0.6)
        pyautogui.hotkey('ctrl', 'f')
        time.sleep(0.6)
        
        pyautogui.typewrite('Number of vCPUs')
        time.sleep(0.6)
        
        pyautogui.press('esc')
        time.sleep(0.6)
        
        actions.send_keys(Keys.ENTER).perform()
        for _ in range(3):
            actions.send_keys(Keys.TAB).perform()
            time.sleep(0.2)
            
        actions.send_keys(Keys.ENTER).perform()
        time.sleep(0.6)
        actions.send_keys(Keys.BACKSPACE).perform()
        time.sleep(0.3)
        actions.send_keys(Keys.BACKSPACE).perform()
        time.sleep(0.3)
        actions.send_keys(Keys.BACKSPACE).perform()
        time.sleep(0.3)
        actions.send_keys(Keys.BACKSPACE).perform()
        time.sleep(0.6)

        actions.send_keys(Keys.BACKSPACE).perform()
        
        time.sleep(0.6)
        actions.send_keys(vcpu).perform()
        print(f"the vcpu is {vcpu}")
        time.sleep(0.6)
        
        
        print("Vcpu got selected")
    
    if ram==0:
        pass
    
    else:
        pyautogui.hotkey('ctrl', 'f')
        time.sleep(0.6)
        
        pyautogui.typewrite('Amount of memory')
        time.sleep(0.6)
        
        pyautogui.press('esc')
        time.sleep(0.6)
        
        actions.send_keys(Keys.ENTER).perform()
        
        time.sleep(0.6)
        for _ in range(3):
            actions.send_keys(Keys.TAB).perform()
            time.sleep(0.2)
            
    
        time.sleep(0.3)
        actions.send_keys(Keys.ENTER).perform()
        time.sleep(0.3)
        actions.send_keys(Keys.BACKSPACE).perform()
        time.sleep(0.3)
        actions.send_keys(Keys.BACKSPACE).perform()
        time.sleep(0.3)
        actions.send_keys(Keys.BACKSPACE).perform()
        time.sleep(0.3)
        actions.send_keys(Keys.BACKSPACE).perform() 
        time.sleep(0.3)
        actions.send_keys(Keys.BACKSPACE).perform()
        time.sleep(0.3)
        actions.send_keys(Keys.BACKSPACE).perform() 
        time.sleep(0.3)
        actions.send_keys(Keys.BACKSPACE).perform()
        time.sleep(0.3)
        actions.send_keys(Keys.BACKSPACE).perform() 
        time.sleep(0.3)



        pyautogui.write(str(ram), interval=0.1)
        pyautogui.press("enter")
        print("ram is selected")

    
    print("✅ vpcu and ram selected")
        
        
        
    

def enable_high_availability(driver,actions,availability):
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.6)
    
    pyautogui.typewrite('Enable High Availability configuration')
    time.sleep(0.6)
    
    pyautogui.press('esc')
    time.sleep(0.6)
    
    actions.key_down(Keys.SHIFT).send_keys(Keys.TAB).key_up(Keys.SHIFT).perform()
    time.sleep(0.6)
    actions.send_keys(Keys.ENTER).perform()
    print("✅ Advanced settings turned on")


def handle_Storage(driver,actions,size):
    time.sleep(0.6)
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.6)
    
    pyautogui.typewrite('Storage (Provisioned Amount)')
    time.sleep(0.6)
    
    pyautogui.press('esc')
    time.sleep(0.6)
    
    actions.send_keys(Keys.ENTER).perform()
    time.sleep(0.6)
    
    for _ in range(3):
        actions.send_keys(Keys.TAB).perform()
        time.sleep(0.2)
    size=int(size)
    actions.send_keys(size).perform()
    time.sleep(0.6)
    for _ in range(2):
        actions.send_keys(Keys.TAB).perform()
        time.sleep(0.2)
    print("✅ Storage selected")




def handle_storage_type(driver,actions):
    time.sleep(0.6)
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.6)
    
    pyautogui.typewrite('Storage Type')
    time.sleep(0.6)
    
    pyautogui.press('esc')
    time.sleep(0.6)
    
    actions.send_keys(Keys.ENTER).perform()
    time.sleep(0.6)
    
    for _ in range(2):
        actions.send_keys(Keys.TAB).perform()
        time.sleep(0.2)
    
    actions.send_keys(Keys.ARROW_RIGHT).perform()
    
    print("✅ Storage type selected")
    


def backup_size(driver,actions,backup_size_value):
    time.sleep(0.6)
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.6)
    
    pyautogui.typewrite('Backup size')
    time.sleep(0.6)
    
    pyautogui.press('esc')
    time.sleep(0.6)
    
    actions.send_keys(Keys.ENTER).perform()
    time.sleep(0.6)
    
    for _ in range(2):
        actions.send_keys(Keys.TAB).perform()
        time.sleep(0.2)
    
    actions.send_keys(backup_size_value).perform()
    time.sleep(0.6)
    for _ in range(2):
        actions.send_keys(Keys.TAB).perform()
        time.sleep(0.2)
    print("✅ Backup size selected")




def get_price_with_js(driver):
    
    try:
        js_script = """
        const element = document.querySelector('span.MyvX5d.D0aEmf');
        return element ? element.textContent.trim() : null;
        """
        price_text = driver.execute_script(js_script)
        
        if price_text and price_text.startswith("$"):
            print("✅ price extracted")
            return price_text
        elif price_text:
            print("❌ Invalid price format")
            return "Invalid price format"
        else:
            print("❌ Price element not found")
            return "Price element not found"
    
    except JavascriptException as e:
        return f"An unexpected JavaScript error occurred: {str(e)}"
    

def add_to_estimate(driver,actions):
    time.sleep(0.6)
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.6)
    
    pyautogui.typewrite('Add to estimate')
    time.sleep(0.6)
    
    pyautogui.press('esc')
    time.sleep(0.6)
    
    actions.send_keys(Keys.ENTER).perform()
    time.sleep(0.6)
    div_element = driver.find_element(By.XPATH, "//div[@class='d5NbRd-EScbFb-JIbuQc PtwYlf' and @data-service-form='7']")
    actions.move_to_element(div_element).click().perform()
    time.sleep(2)
   
    
    print("✅ Added to estimate")

#========================================================================================================#
def sud_pricing(driver,actions,service_type_value,region,cloud_sql_edition_value,instance_value,usage_time_value,instance_type,HA,storage_type,size,backup_size_value,vcpu,ram):
    service_type(driver,actions,service_type_value)
    select_region(driver,actions,region)
    advanced_toggle_on(driver,actions)
    cloud_sql_edition(driver,actions,cloud_sql_edition_value)
    #Specify_usage_time(driver,actions)
    instance(driver,actions,instance_value)
    time.sleep(0.3)
    usage_time(driver,actions,usage_time_value)
    time.sleep(0.3)
    select_sql_instance_type(driver,actions,instance_type)
    time.sleep(0.3)
    if instance_type.lower()!= "db-f1-micro" or instance_type.lower()!= "db-g1-small":
        vcpu_handle(driver,actions,vcpu,ram)
    if HA.upper()=="HA":
        enable_high_availability(driver,actions,HA)
    handle_Storage(driver,actions,size)
    if storage_type=="HDD":
        handle_storage_type(driver,actions) 
    
    backup_size(driver,actions,backup_size_value)
    time.sleep(10)
    
    current_url = driver.current_url
    time.sleep(10)

    price=get_price_with_js(driver)
    
    print(price,current_url)
    
    
    print("✅ SUD pricing done")
    
    return price, current_url

def one_year_pricing(driver,actions,service_type_value,region,cloud_sql_edition_value,instance_value,usage_time_value,instance_type,HA,storage_type,size,backup_size_value,vcpu,ram):
    service_type(driver,actions,service_type_value)
    select_region(driver,actions,region)
    advanced_toggle_on(driver,actions)
    cloud_sql_edition(driver,actions,cloud_sql_edition_value)
    #Specify_usage_time(driver,actions)
    instance(driver,actions,instance_value)
    usage_time(driver,actions,usage_time_value)
    time.sleep(0.3)
    select_sql_instance_type(driver,actions,instance_type)
    time.sleep(0.3)
    if instance_type.lower()!= "db-f1-micro" or instance_type.lower()!= "db-g1-small":
        vcpu_handle(driver,actions,vcpu,ram)
    if HA=="HA":
        enable_high_availability(driver,actions,HA)
    handle_Storage(driver,actions,size)
    if storage_type=="HDD":
        handle_storage_type(driver,actions) 
    
    backup_size(driver,actions,backup_size_value)
    
    time.sleep(0.6)
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.6)
    
    pyautogui.typewrite('Committed use discount options')
    time.sleep(0.6)
    
    pyautogui.press('esc')
    time.sleep(0.6)
    
    actions.send_keys(Keys.ENTER).perform()
    time.sleep(0.6)
    
    for _ in range(2):
        actions.send_keys(Keys.TAB).perform()
    
    actions.send_keys(Keys.ARROW_RIGHT).perform()
    time.sleep(0.4)
    actions.send_keys(Keys.ENTER).perform()
    current_url = driver.current_url
    time.sleep(10)

    price=get_price_with_js(driver)
    
    print(price,current_url)
    print("✅ One year pricing selected")
    return price,current_url

def three_year_pricing(driver,actions,service_type_value,region,cloud_sql_edition_value,instance_value,usage_time_value,instance_type,HA,storage_type,size,backup_size_value,vcpu,ram):
    service_type(driver,actions,service_type_value)
    select_region(driver,actions,region)
    advanced_toggle_on(driver,actions)
    cloud_sql_edition(driver,actions,cloud_sql_edition_value)
    #Specify_usage_time(driver,actions)
    instance(driver,actions,instance_value)
    time.sleep(0.3)
    usage_time(driver,actions,usage_time_value)
    time.sleep(0.3)
    select_sql_instance_type(driver,actions,instance_type)
    time.sleep(5)
    if instance_type.lower()!= "db-f1-micro" or instance_type.lower()!= "db-g1-small":
        vcpu_handle(driver,actions,vcpu,ram)
    if HA=="HA":
        enable_high_availability(driver,actions,HA)
    handle_Storage(driver,actions,size)
    if storage_type=="HDD":
        handle_storage_type(driver,actions) 
    
    backup_size(driver,actions,backup_size_value)
    
    
    
    
    
    time.sleep(0.6)
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.6)
    
    pyautogui.typewrite('Committed use discount options')
    time.sleep(0.6)
    
    pyautogui.press('esc')
    time.sleep(0.6)
    
    actions.send_keys(Keys.ENTER).perform()
    time.sleep(0.6)
    
    for _ in range(2):
        actions.send_keys(Keys.TAB).perform()
    
    actions.send_keys(Keys.ARROW_RIGHT).perform()
    time.sleep(0.4)
    actions.send_keys(Keys.ARROW_RIGHT).perform()
    time.sleep(0.4)
    actions.send_keys(Keys.ENTER).perform()
    current_url = driver.current_url
    time.sleep(10)
    price=get_price_with_js(driver)
    
    print(price,current_url)
    print("✅ Three year pricing selected")
    return current_url, price


#========================================================================================================#


def read_input_values(file_path):# Read the input sheet
    if file_path.endswith(".csv"):
        df = pd.read_csv(file_path)
    else:
        df = pd.read_excel(file_path)
        
    df.fillna({
        "SQL Type": "MySQL",
        "Datacenter Location": "us-central1",
        "Cloud SQL": "Enterprise",
        "No. of Instances": 1,
        "Avg no. of hrs": 730,
        "Instance Type": "db-n1-standard-2",
        "HA/Non-HA": "Non-HA",
        "Disk Type": "HDD",
        "Storage Amt": 256,
        "vCPUs":0,
        "RAM":0,
    }, inplace=True)
    
    return df

def save_to_excel(data, filename):
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)
def extract_sheet_id(sheet_url):
    pattern = r"https://docs\.google\.com/spreadsheets/d/([a-zA-Z0-9-_]+)"
    match = re.search(pattern, sheet_url)
    if match:
        return match.group(1)
    else:
        raise ValueError("Invalid Google Sheet URL")

def download_sheet(sheet_url):
        try:
            print("downloading the sheet !!")
            sheet_id = extract_sheet_id(sheet_url)
            csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
            response = requests.get(csv_url)

            if response.status_code == 200:
                with open("sheet.csv", "wb") as f:
                    f.write(response.content)
                print("Google Sheet downloaded as sheet.csv")
            else:
                print("Failed to download sheet. HTTP Status Code:", response.status_code)
        except ValueError as e:
            print(e)
        except Exception as e:
            print("An error occurred:", e)


    
def setup_driver():
    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": os.path.join(os.getcwd(), "downloads"),
        "download.prompt_for_download": False,
        "safebrowsing.enabled": True,
    }
    chrome_options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window()
    driver.get("https://cloud.google.com/products/calculator")
    driver.implicitly_wait(10)
    return driver

def main():
    file_path="jane new GcpData- - CloudSql.csv"
    df = read_input_values(file_path) 
    results = [{
        "SL No": index + 1,
        "Instance": str(row["No. of Instances"]),
        "Cloud SQL Edition": row["Cloud SQL "],
        "SUD Price": "",
        "SUD URL": "",
        "One Year Price": "",
        "One Year URL": "",
        "Three Year Price": "",
        "Three Year URL": ""
    } for index, row in df.iterrows()]

    #sud pricing
    driver = setup_driver()
    actions = ActionChains(driver)
    for index, row in df.iterrows():
        if index == 0:
            home_page(driver,actions)
        
        
        sud_price, sud_current_url = sud_pricing(driver, actions, row["SQL Type"], row["Datacenter Location"], row["Cloud SQL "],float(row["No. of Instances"]), int(row["Avg no. of hrs"]), str(row["Instance Type"]),row["HA/Non-HA"], row["Disk Type"], int(row["Storage Amt"]),int(row["Backup"]),int(row["vCPUs"]),int(row["RAM"]))
        
        results[index]["SUD Price"] = sud_price
        results[index]["SUD URL"] = sud_current_url
        if index < len(df) - 1:
            add_to_estimate(driver,actions)
            
    driver.quit()
    
    
    
    # Processing One-Year Pricing
    driver = setup_driver()
    actions = ActionChains(driver)
    for index, row in df.iterrows():
        if index == 0:
            home_page(driver,actions)
        one_year_price, one_year_current_url = one_year_pricing(driver, actions, row["SQL Type"], row["Datacenter Location"], row["Cloud SQL "],float(row["No. of Instances"]), int(row["Avg no. of hrs"]), str(row["Instance Type"]),row["HA/Non-HA"], row["Disk Type"], int(row["Storage Amt"]),int(row["Backup"]),int(row["vCPUs"]),int(row["RAM"]))
        results[index]["One Year Price"] = one_year_price
        results[index]["One Year URL"] = one_year_current_url
        if index < len(df) - 1:
            add_to_estimate(driver,actions)
    driver.quit()
    
    
    # Processing Three-Year Pricing
    driver = setup_driver()
    actions = ActionChains(driver)
    
    
    for index, row in df.iterrows():
        if index == 0:
            home_page(driver,actions)
        three_year_price, three_year_current_url = three_year_pricing(driver, actions, row["SQL Type"], row["Datacenter Location"], row["Cloud SQL "],float(row["No. of Instances"]), int(row["Avg no. of hrs"]), str(row["Instance Type"]),row["HA/Non-HA"], row["Disk Type"], int(row["Storage Amt"]),int(row["Backup"]),int(row["vCPUs"]),int(row["RAM"]))
        results[index]["Three Year Price"] = three_year_current_url
        results[index]["Three Year URL"] = three_year_price
        if index < len(df) - 1:
            add_to_estimate(driver,actions)
    driver.quit()
    
    save_to_excel(results, "pricing_summary.xlsx")
    print("✅ All pricing done and saved in pricing_summary.xlsx")



if __name__ == "__main__":
    start_time=time.time()
    main()
    stop_time=time.time()
    print(f"Time taken is {stop_time-start_time}")