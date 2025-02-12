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



app=Flask(__name__)

process_status = {}


index_file = "index.json"



input_file = "sheet.csv"  # Replace with your input file path


output_file_filtered = "output.csv" 
input_filtered_file="output.csv"


output_file = "output_results.csv"

sender_email = "srujan.int@niveussolutions.com"
sender_password = "rmlh ikej rtmz ejme"
subject = "Compute Calculation Results"
body = "Please find the attached file for the results of the computation."
file_path = "output_results.xlsx"






#=============================================================================================#
def get_on_demand_pricing( os_name, no_of_instances,hours_per_day, machine_family, series, machine_type, vCPU, ram, boot_disk_capacity, region,machine_class):
    pass
    
   
def get_sud_pricing( os_name, no_of_instances,hours_per_day, machine_family, series, machine_type, vCPU, ram, boot_disk_capacity, region,machine_class):
    pass
    
    
def get_one_year_pricing(os_name, no_of_instances,hours_per_day, machine_family, series, machine_type, vCPU, ram, boot_disk_capacity, region,machine_class):
    pass

def  get_three_year_pricing(os_name, no_of_instances,hours_per_day, machine_family, series, machine_type, vCPU, ram, boot_disk_capacity, region,machine_class):
    pass


#=============================================================================================#


def main(sheet_url,recipient_email):
    download_sheet(sheet_url)

    process_csv(input_file, output_file_filtered)
    sheet = pd.read_csv(input_filtered_file)
    results = []

    for index, row in sheet.iterrows():
        os_name = row["OS with version"]
        print(os_name)
        no_of_instances = round(float(row["No. of Instances"]), 2) if pd.notna(row["No. of Instances"]) else 0.00
        machine_family = row["Machine Family"].lower() if pd.notna(row["Machine Family"]) else "general purpose"
        series = row["Series"].upper() if pd.notna(row["Series"]) else "E2"
        machine_type = row["Machine Type"].lower() if pd.notna(row["Machine Type"]) else "custom"
        vCPU = row["vCPUs"] if pd.notna(row["vCPUs"]) else 0
        ram = row["RAM"] if pd.notna(row["RAM"]) else 0
        boot_disk_capacity = row["BootDisk Capacity"] if pd.notna(row["BootDisk Capacity"]) else 0
        region = row["Datacenter Location"] if pd.notna(row["Datacenter Location"]) else "Mumbai"
        hours_per_day = int(row["Avg no. of hrs"]) if pd.notna(row["Avg no. of hrs"]) else 730
        machine_class = str(row["Machine Class"]) if pd.notna(row["Machine Class"]) else "regular"
        #print(hours_per_day)
        machine_class=machine_class.lower()
        
       
        
        print(f"Processing row {index + 1} with OS: {os_name}, Instances: {no_of_instances}, machine family: {machine_family}, series: {series}, machine type: {machine_type}")

        row_result = {
            "Row Index": index + 1,
            "OS with version": os_name,
            "No. of Instances": no_of_instances,
            "Machine Family": machine_family,
            "On-Demand URL": None,
            "On-Demand Price": None,
            "SUD URL": None,
            "SUD Price": None,
            "1-Year URL": None,
            "1-Year Price": None,
            "3-Year URL": None,
            "3-Year Price": None
        }

        for iteration in range(4):  
            try:
                
                
                
                if  machine_class=="preemptible": #if it is less than on demand price  ///////even if it greater then 730 and (spot/premtible eny condition the price is ondemand)
                    if iteration==0:
                        print(f"Iteration {iteration + 1}: Getting on-demand price and link (e2-micro)")
                        row_result["On-Demand URL"], row_result["On-Demand Price"] = get_on_demand_pricing(
                            os_name, no_of_instances, hours_per_day, machine_family, series, machine_type, vCPU, ram, boot_disk_capacity, region,machine_class
                        )

                    if iteration==1:
                        
                        print(f"Iteration {iteration + 1}: Getting sustained use discount (SUD) price and link")
                        row_result["SUD URL"], row_result["SUD Price"] = row_result["On-Demand URL"], row_result["On-Demand Price"]
                        
                        row_result["1-Year URL"], row_result["1-Year Price"] = row_result["SUD URL"], row_result["SUD Price"]
                        row_result["3-Year URL"], row_result["3-Year Price"] = row_result["SUD URL"], row_result["SUD Price"]
                        break
                    
                if  machine_class=="regular" and hours_per_day < 730:
                    if iteration==0:
                        print(f"Iteration {iteration + 1}: Getting on-demand price and link (e2-micro)")
                        row_result["On-Demand URL"], row_result["On-Demand Price"] = get_on_demand_pricing(
                            os_name, no_of_instances, hours_per_day, machine_family, series, machine_type, vCPU, ram, boot_disk_capacity, region,machine_class
                        )

                    if iteration==1:
                        
                        print(f"Iteration {iteration + 1}: Getting sustained use discount (SUD) price and link")
                        row_result["SUD URL"], row_result["SUD Price"] = get_sud_pricing(
                            os_name, no_of_instances, hours_per_day, machine_family, series, machine_type, vCPU, ram, boot_disk_capacity, region,machine_class
                        )
                        
                        row_result["1-Year URL"], row_result["1-Year Price"] = row_result["SUD URL"], row_result["SUD Price"]
                        row_result["3-Year URL"], row_result["3-Year Price"] = row_result["SUD URL"], row_result["SUD Price"]
                        break
                
                
                else:
                    if iteration == 0:
                        print(f"Iteration {iteration + 1}: Getting on-demand price and link")
                        row_result["On-Demand URL"], row_result["On-Demand Price"] = get_on_demand_pricing(
                            os_name, no_of_instances, hours_per_day, machine_family, series, machine_type, vCPU, ram, boot_disk_capacity, region,machine_class
                        )
                    elif iteration == 1:
                        if series=="E2":
                            print(f"Iteration {iteration + 1}: Getting sustained use discount (SUD) price and link")
                            row_result["SUD URL"], row_result["SUD Price"] = row_result["On-Demand URL"], row_result["On-Demand Price"] 
                            
                        
                        else:
                            print(f"Iteration {iteration + 1}: Getting sustained use discount (SUD) price and link")
                            row_result["SUD URL"], row_result["SUD Price"] = get_sud_pricing(
                                os_name, no_of_instances, hours_per_day, machine_family, series, machine_type, vCPU, ram, boot_disk_capacity, region,machine_class
                            )
                    elif iteration == 2:
                        print(f"Iteration {iteration + 1}: Getting 1-year commitment price and link")
                        row_result["1-Year URL"], row_result["1-Year Price"] = get_one_year_pricing(
                            os_name, no_of_instances, hours_per_day, machine_family, series, machine_type, vCPU, ram, boot_disk_capacity, region,machine_class
                        )
                    elif iteration == 3:
                        print(f"Iteration {iteration + 1}: Getting 3-year commitment price and link")
                        row_result["3-Year URL"], row_result["3-Year Price"] = get_three_year_pricing(
                            os_name, no_of_instances, hours_per_day, machine_family, series, machine_type, vCPU, ram, boot_disk_capacity, region,machine_class
                        )
            except Exception as e:
                print(f"Error processing row {index + 1}, iteration {iteration + 1}: {e}")
                if iteration == 0:
                    row_result["On-Demand URL"] = "Error"
                    row_result["On-Demand Price"] = "Error"
                elif iteration == 1:
                    row_result["SUD URL"] = "Error"
                    row_result["SUD Price"] = "Error"
                elif iteration == 2:
                    row_result["1-Year URL"] = "Error"
                    row_result["1-Year Price"] = "Error"
                elif iteration == 3:
                    row_result["3-Year URL"] = "Error"
                    row_result["3-Year Price"] = "Error"
            finally:
                time.sleep(5)

        results.append(row_result)


    output_file = "output_results.xlsx"  # Excel file extension

    output_df = pd.DataFrame(results)
    output_df.to_excel(output_file, index=False)
    print(f"Results saved to {output_file}")
    print("sending mail!!")
    send_email_with_attachment(sender_email, sender_password, recipient_email, subject, body, file_path)


@app.route('/calculate',methods=["POST"])
def run_automation():
    sheet = request.form.get('sheet')
    email = request.form.get('email')
    process_status[email] = "Processing"
    main(sheet,email)
    process_status[email] = "Completed"
    return "process completed sucessfully"




if __name__ == "__main__":
    
    app.run(debug=True,use_reloader=False,host='0.0.0.0')
