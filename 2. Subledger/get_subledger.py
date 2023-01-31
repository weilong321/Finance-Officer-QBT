from selenium.webdriver.chrome.webdriver import WebDriver as Chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from time import sleep

from openpyxl import load_workbook
from openpyxl import utils
import pandas as pd
import numpy as np

import os
import shutil

# Logging into AGM
APXNZ = "https://www.agencymanager.amadeus.com/agm-connect-1.2.40/?configfile=/ahpapxnz1e0af369"
USER_XPATH = "//*[@id='UserPrf']"
PASS_XPATH = "//*[@id='Password']"
LOGON_XPATH = "//*[@id='Logon']"
LOGON_TOKEN_XPATH = "//*[@id='edtToken']"
REQUEST_NEW_TOKEN_XPATH = "//*[@id='btnRequest']"
OK_XPATH = "//*[@id='OK']"
USERNAME = "user"
PASSWORD = "password"

# Inside APX NZ AGM (successful login)
MENU_XPATH = "//*[@id='tnFavMenu']/fieldset/agm-tab-bar/ul/div/div[2]/agm-tab-button"
BACK_OFFICE_XPATH = "//*[@id='MI222']/button"
OPEN_ITEMS_XPATH = "//*[@id='MI317']/button"
OPEN_ITEMS_CONSULTATION_XPATH = "//*[@id='MI318']/button"
GL_ACCOUNT_XPATH = "//*[@id='AccCde']"
SL_ACCOUNT_XPATH = "//*[@id='SLAcc']"
EXPORT_EXISTING_TO_CSV = "//*[@id='brw_1']/div[2]/agm-grid-button-icon[3]/button"
EXPORT_ALL_TO_CSV = "//*[@id='brw_1']/div[2]/agm-grid-button-icon[4]/button"
CONFIRM_EXPORT_ALL = "/html/body/ngb-modal-window/div/div/agm-message-box-content/div[3]/button[1]"
ALL_DATA_BUTTON = "//*[@id='brw_1']/div[1]/table/thead/tr/th[1]"
EXPORT_CANCEL_BUTTON = "/html/body/agm-root/div/agm-export-progress/div[1]/div/div/div[3]/button"
EXPORT_POPUP = "/html/body/agm-root/div/agm-export-progress/div[1]"


def get_df_ledgers(path):
    df = pd.read_excel(path, sheet_name="ITBR Summary")
    GL_account = df.iat[0, 0]
    df = df.drop(df.index[:2])
    df.columns = df.iloc[0]
    df = df.drop(df.index[0])
    df.columns = ["Supplier No", "Supplier Name", "Currency", "Sep and prior", "Oct-22", "Nov-22", "Dec-22", "Total"]
    df = df.iloc[:-5].reset_index(drop=True).fillna(0)
    df = df[(df["Sep and prior"] != 0) & (df["Supplier Name"] != "IATA BSP")].reset_index(drop=True)
    SL_account_list = df[df['Sep and prior'] != 0]['Supplier No'].tolist()
    return GL_account, SL_account_list, df

def get_style(xpath):
    wait = WebDriverWait(driver, 10)
    element = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
    element = driver.find_element(By.XPATH, xpath)
    style = driver.execute_script("return arguments[0].getAttribute('style');", element)
    return style

def find_type_click_element(xpath, not_button=None, input=None, needs_enter=None, csv=None, wait=True, style=None):
    if wait:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
    element = driver.find_element(By.XPATH, xpath)
    if csv:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        action = ActionChains(driver)
        action.move_to_element(element).perform()
        sleep(1)
    element.click()
    # if style:
    #     style = driver.execute_script("return arguments[0].getAttribute('style');", element)
    #     if style != 'color: rgb(0, 133, 64);':
    #         find_type_click_element(EXPORT_ALL_TO_CSV)
    #     print(style)
    if not_button:
        element.clear()
        element.send_keys(input)
    if needs_enter:
        element.send_keys(Keys.ENTER)
    return

def APXNZ_login():
    find_type_click_element(USER_XPATH, not_button=True, input=USERNAME)
    find_type_click_element(PASS_XPATH, not_button=True, input=PASSWORD)
    find_type_click_element(LOGON_XPATH)
    return

def get_token():
    find_type_click_element(REQUEST_NEW_TOKEN_XPATH)
    logon_key = input("Type in logon token here: \n")
    find_type_click_element(LOGON_TOKEN_XPATH, not_button=True, input=logon_key)
    find_type_click_element(OK_XPATH)
    return

def rename_move_file(general, subledger, is_first):
    if is_first:
        is_first = False
        sleep(2)
    # desktop_path = "C:/Users/ctm_mchen/OneDrive - Helloworld Travel Ltd/Desktop/General_Ledger_{}/".format(general)
    desktop_path = "C:/Users/Matthew Chen/Desktop/subledger/General_Ledger_{}/".format(general)
    if not os.path.exists(desktop_path):
        os.mkdir(desktop_path)
    # initial_path = "C:/Users/ctm_mchen/Downloads"
    initial_path = "C:/Users/Matthew Chen/Downloads"
    filename = max([initial_path + "/" + f for f in os.listdir(initial_path)],key=os.path.getctime)
    ext = os.path.splitext(filename)[1]
    shutil.move(filename,os.path.join(desktop_path,"Subledger_{}{}".format(subledger, ext)))

def get_subledger(gl, sls):
    is_first = True
    find_type_click_element(MENU_XPATH)
    find_type_click_element(BACK_OFFICE_XPATH)
    find_type_click_element(OPEN_ITEMS_XPATH)
    find_type_click_element(OPEN_ITEMS_CONSULTATION_XPATH)
    find_type_click_element(GL_ACCOUNT_XPATH, not_button=True, input=gl)
    for subledger in sls:
        find_type_click_element(SL_ACCOUNT_XPATH, not_button=True, input=subledger, needs_enter=True)
        sleep(1)
        if get_style(ALL_DATA_BUTTON) != 'color: rgb(0, 133, 64);':
            find_type_click_element(EXPORT_ALL_TO_CSV, csv=True)
            find_type_click_element(CONFIRM_EXPORT_ALL)
            sleep(1)
            wait = WebDriverWait(driver, 120)
            wait.until(EC.invisibility_of_element_located((By.XPATH, EXPORT_POPUP)))
        else:
            find_type_click_element(EXPORT_EXISTING_TO_CSV, csv=True)
        print("downloaded {}".format(subledger))
        rename_move_file(gl, subledger, is_first)
        print("moved {} file".format(subledger))
    return

def login_download_csv(general, subledgers):
    global driver
    driver = Chrome("chromedriver.exe")
    driver.maximize_window()
    driver.get(APXNZ)
    APXNZ_login()
    # get_token()
    get_subledger(general, subledgers)
    sleep(2)
    return

def convert_to_numeric(val):
    try:
        return pd.to_numeric(val)
    except ValueError:
        return val.replace(',', '')

def read_subledger_csv(path):
    ls = [0, 0, 0, 0, 0, 0, 0, 0, 0]
    temp = pd.read_csv(path, skiprows = 1)
    temp = temp.applymap(convert_to_numeric)
    if temp.shape[0] == 0:
        return ls
    temp["Amount local"] = pd.to_numeric(temp["Amount local"], errors='coerce')
    table = pd.pivot_table(temp, values='Amount local', index='Entry period', aggfunc=np.sum)
    for period in table.index:
        fy = str(period)
        amount = round(table.loc[period, 'Amount local'], 2)
        if fy.startswith("2017"):
            ls[0] += amount
            ls[8] += amount
        elif fy.startswith("2018"):
            ls[1] += amount
            ls[8] += amount
        elif fy.startswith("2019"):
            ls[2] += amount
            ls[8] += amount
        elif fy.startswith("2020"):
            ls[3] += amount
            ls[8] += amount
        elif fy.startswith("2021"):
            ls[4] += amount
            ls[8] += amount
        elif fy.startswith("20220"):
            ls[5] += amount
            ls[8] += amount
        elif fy.startswith("20221"):
            ls[6] += amount
            ls[8] += amount
        elif fy == "202301" or fy == "202302" or fy == "202303":
            ls[7] += amount
            ls[8] += amount
    return ls

def get_financial_year_sum(main_df, general):
    folder_path = "General_Ledger_{}".format(general)
    files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    files = sorted(files, key=lambda x: len(os.path.basename(x)))
    ls = []
    for file in files:
        temp = read_subledger_csv(folder_path + '/' + file)
        ls.append(temp)
    fy_df = pd.DataFrame(ls, columns=['FY17', 'FY18', 'FY19', 'FY20', 'FY21', 'FY22 (to Mar22)', 'FY22 (Apr-Jun22)', 'FY22 (Jul-Sep22)', 'Total (Sep and prior)'])
    combined = pd.concat([main_df[['Supplier No', 'Sep and prior']], fy_df], axis=1).fillna(0)
    combined = combined.assign(Diff = combined['Sep and prior'] - combined['Total (Sep and prior)'])
    return combined

def add_df_to_wb(combined, path):
    wb = load_workbook(path)
    ws = wb.create_sheet("Combined Summary")
    for col_idx, value in enumerate(combined.columns.tolist()):
        ws.cell(row=1, column=col_idx+1).value = value
    for row_idx, row in combined.iterrows():
        for col_idx, cell_value in enumerate(row.tolist()):
            try:
                rounded_value = round(float(cell_value), 2)
                ws.cell(row=row_idx+2, column=col_idx+1).value = rounded_value
            except:
                ws.cell(row=row_idx+2, column=col_idx+1).value = cell_value
        ws.freeze_panes = 'A2'
        for column in ws.columns:
            for cell in column:
                column_letter = utils.get_column_letter(cell.column)
                ws.column_dimensions[column_letter].auto_size = True
    wb.save(path)
    return



#####################################################################
#                                                                   #  
#                             MAIN                                  #
#                                                                   #
#####################################################################

path = "APX ITBR Aging - Dec22 (TESTING).xlsx"
general, subledgers, main_df = get_df_ledgers(path)

# read_subledger_csv("General_Ledger_5080020000/Subledger_1042.csv")

login_download_csv(general, subledgers)

combined = get_financial_year_sum(main_df, general)

add_df_to_wb(combined, path)