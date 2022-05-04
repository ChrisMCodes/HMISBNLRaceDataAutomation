#!/usr/bin/env python

from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from time import sleep
from datetime import date
import os
import openpyxl as op
import warnings

# get login info
username = input("Enter your HMIS username: ")
password = input("Enter your password: ")

# Functions to open and retrieve BNL data

def get_workbook():
    # Retrieving the bnl
    all_files = os.listdir()
    list_xlsx = [file for file in all_files if "BNL" in file]
    bnl = list_xlsx[0]
    return bnl

def get_data(bnl):
    # Saving as an array of rows
    workbook = op.load_workbook(filename=bnl)
    ws = workbook['Details']
    all_cells = [[cell.value for cell in row] for row in ws.iter_rows()]
    data_only = [subl for subl in all_cells if subl[0] != "HMIS ID" and subl[25] == "Yavapai"]
    return data_only


def login_hmis(username, password):
    # Logs into site
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        browser = webdriver.Chrome(ChromeDriverManager().install())
        browser.get("https://ruralaz.servicept.com/")
        user_input = browser.find_element_by_id('formfield-login')
        user_input.send_keys(username)
        password_input = browser.find_element_by_id('formfield-password')
        password_input.send_keys(password)
        submit_button = browser.find_element_by_id("LoginView.fbtn_submit")
        sleep(1)
        submit_button.click()
        sleep(3)
    return browser

def set_eda(browser):
    # Sets EDA to Coordinated Entry
    eda = browser.find_element_by_xpath("//*[@id=\"enterDataAsModeAnchor\"]/a")
    eda.click()
    sleep(2)
    provider = browser.find_element_by_xpath(
    "/html/body/div[6]/div/table/tbody/tr[2]/td[2]/table/tbody/tr[3]/td/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/input")
    provider.send_keys("717")
    sleep(1)
    submit_button = browser.find_element_by_xpath(
    "/html/body/div[6]/div/table/tbody/tr[2]/td[2]/table/tbody/tr[3]/td/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td[3]/div")
    submit_button.click()
    sleep(1)
    add_eda = browser.find_element_by_xpath(
    "/html/body/div[6]/div/table/tbody/tr[2]/td[2]/table/tbody/tr[3]/td/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[4]/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[1]")
    add_eda.click()
    sleep(1)

def get_ids(data):
    # Extracts ids only from BNL and returns list
    counter = 0
    all_ids = []
    for i in data:
        all_ids.append(data[counter][0])
        counter += 1
    return all_ids


def get_amount_of_time_homeless(data):
    counter = 0
    all_aoth = []
    for i in data:
        # This if is necessary to escape NoneType values
        if data[counter][16]:
            diff_days = date.today() - data[counter][16].date()
            num_days = int(diff_days.days)
        else:
            num_days = 0
        all_aoth.append(num_days)
        counter += 1
    return all_aoth

def get_disability_info(data):
    counter = 0
    disability = []
    for i in data:
        disability.append(1 if data[counter][15] == "Yes (HUD)" else 0)
        counter += 1
    return disability

def get_smi_info(data):
    counter = 0
    smi = []
    for i in data:
        smi.append(1 if data[counter][8] == "Yes" else 0)
        counter += 1
    return smi

def get_dv_info(data):
    counter = 0
    dv = []
    for i in data:
        dv.append(1 if data[counter][29] == "Yes (HUD)" else 0)
        counter += 1
    return dv


def main(username, password):
    print("Working...")
    bnl = get_workbook()
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        data = get_data(bnl)
    ids = get_ids(data)
    aoth = get_amount_of_time_homeless(data)
    disability = get_disability_info(data)
    smi = get_smi_info(data)
    dv = get_dv_info(data)

    # Write CSV with data above
    with open("bnl_race_data_output.csv", "w") as bnlrd:
        bnlrd.write("Client HMIS ID, Days homeless, Disability (1 = yes, 0 = no), SMI (1 = yes, 0 = no), DV (1 = yes, 0 = no), Primary Race, Ethnicity\n")
        browser = login_hmis(username, password)
        set_eda(browser)
        # search for clients and go to bos assessment
        counter = 0
        for id in ids:
            bnlrd.write(str(id) + "," + str(aoth[counter]) + "," + str(disability[counter]) + "," +
                        str(smi[counter]) + "," + str(dv[counter]) + ",")
            client_point = browser.find_element_by_id("navigation-link.clientpt")
            client_point.click()
            sleep(1.5)

            search = browser.find_element_by_id("ClientSearchView.clientId-textbox")
            search.send_keys(id)

            button = browser.find_element_by_id("ClientSearchView.clientIdSubmitBtn")
            button.click()

            sleep(1.5)

                # extract client's primary race

            try:
                # Switch to client profile view
                browser.find_element_by_xpath("/html/body/table[2]/tbody/tr[4]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[3]/td/table/tbody/tr[1]/td/table/tbody/tr/td[2]/div/div[6]/div").click()
                sleep(1)
                # Get race
                if browser.find_element_by_xpath("/html/body/table[2]/tbody/tr[4]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[4]/td[3]/table/tbody/tr/td[3]"):
                    race = browser.find_element_by_xpath("/html/body/table[2]/tbody/tr[4]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[4]/td[3]/table/tbody/tr/td[3]")
                    write_race = race.text
                    bnlrd.write(write_race)
                    print(write_race)
                    sleep(0.5)
                # Get ethnicity
                if browser.find_element_by_xpath("/html/body/table[2]/tbody/tr[4]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[6]/td[3]/table/tbody/tr/td[3]/div"):
                    ethnicity = browser.find_element_by_xpath("/html/body/table[2]/tbody/tr[4]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[6]/td[3]/table/tbody/tr/td[3]/div")
                    write_ethnicity = ethnicity.text
                    bnlrd.write(write_ethnicity + ",")
                    print(write_ethnicity)
                    sleep(1.5)
            except Exception as e:
                print("An error occurred")
                print()
                print(e)
            finally:
                bnlrd.write("\n")
                counter += 1
        bnlrd.close()
        browser.quit()
        
    

main(username, password)
    
