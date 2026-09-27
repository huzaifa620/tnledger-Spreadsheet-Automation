# Generated from "tn ledger foreclosure.ipynb" so the VA hub can run it outside Jupyter.
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from openpyxl import Workbook
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl.styles import Font, Color, Alignment, Border, Side
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.utils import get_column_letter
import tkinter as tk
import csv, os, sys, time, re, random
import PySimpleGUI as sg
from datetime import date
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2 import service_account

# 4/1/2022
#month/day/year
today = str(date.today()).split('-')
theme_name_list = sg.theme_list()
while True:
    sg.theme(theme_name_list[random.randint(0, len(theme_name_list)-1)])
    #define layout
    layout = [[
        sg.Text('Choose Default Date',
                size=(20, 1),
                font='Ubuntu',
                justification='left')],
              [
                  sg.Input(key='-IN4-', size=(20, 1)),
                  sg.CalendarButton(
                      'Calendar1',
                      font="Ubuntu",
                      target='-IN4-',
                      default_date_m_d_y=(int(today[1]), int(today[2]),
                                          int(today[0])),
                  )
              ],
              [
                  sg.Button('OK', font=('Ubuntu', 12)),
                  sg.Button('CANCEL', font=('Ubuntu', 12))
              ]]
    win = sg.Window('Tnledger', layout)
    e, v = win.read()
    print(e, v)
    if e == None or e == "CANCEL":
        print("exit")
        win.close()
        raise SystemExit(0)
    else:
        if  v['-IN4-'] == None or v['-IN4-'] == '':
            print('Enter the date correctly')
            win.close()
            continue
        else:
            win.close()
            break

currentdate = time.strptime(v['-IN4-'].split(' ')[0], "%Y-%m-%d")

selected_date=f'{currentdate.tm_mon}/{currentdate.tm_mday}/{currentdate.tm_year}'

driver = webdriver.Chrome(ChromeDriverManager().install())
driver.get(f'https://www.tnledger.com/Notices.aspx?noticesDate={selected_date}')
driver.implicitly_wait(30)
date=driver.current_url[driver.current_url.find('=')+1:]
driver1 = webdriver.Chrome(ChromeDriverManager().install())
all_data = []
date=driver.current_url[driver.current_url.find('=')+1:]
forclosure_rows=driver.find_elements(By.XPATH, '//table[@id="ContentPane_ForeclosureGridView"]/tbody/tr')
del forclosure_rows[0]
for i in forclosure_rows:
    individual_rows=i.find_elements(By.TAG_NAME, 'td')
    each_row_data=[]
    for j in (individual_rows):
        if j.text=="View":
            current_link=(j.find_element(By.TAG_NAME, 'a').get_attribute('href'))
            link_id=current_link[current_link.find('(')+1:current_link.find(')')].split(',')[0].replace("'","")
            
            driver1.get(f"https://www.tnledger.com/Search/Details/ViewNotice.aspx?id={link_id}&date={date}")
            driver1.implicitly_wait(30)
            table_data=driver1.find_elements(By.XPATH, '//table[@class="data-table"]/tbody/tr')
            for counter,k in enumerate(table_data):
                horizontal_data=k.find_elements(By.TAG_NAME, 'td')
                check=horizontal_data.pop(0).text
                if  'Original Trustee:'in check:
                    break
                for l in horizontal_data:
                    each_row_data.append(l.text)
    all_data.append(each_row_data)
driver1.close()
        
################################################################################################

# bold_font = Font(bold=True)
# bold_font = Font(bold=True)
# center_aligned_text = Alignment(horizontal="center")
# double_border_side = Side(border_style="double")
# square_border = Border(top=double_border_side,
#                 right=double_border_side,
#                 bottom=double_border_side,
#                 left=double_border_side)
# workbook = Workbook()
# sheet = workbook.active
# sheet["A1"] = "Borrower's Name"
# sheet["A1"].font = bold_font
# sheet["A1"].alignment = center_aligned_text
# sheet["A1"].border = square_border
# sheet["B1"] = "Address"
# sheet["B1"].font = bold_font
# sheet["B1"].alignment = center_aligned_text
# sheet["B1"].border = square_border
# sheet["C1"] = "Remaining Address"
# sheet["C1"].font = bold_font
# sheet["C1"].alignment = center_aligned_text
# sheet["C1"].border = square_border
# dim_holder = DimensionHolder(worksheet=sheet)

# for col in range(sheet.min_column, sheet.max_column + 1):
#     dim_holder[get_column_letter(col)] = ColumnDimension(sheet, min=col, max=col, width=20)

# sheet.column_dimensions = dim_holder
# for p,q in enumerate(all_data):
#     sheet[f"A{p+2}"]=q[0]           
#     sheet[f"B{p+2}"]=q[1]
#     sheet[f"C{p+2}"]=q[2]

# workbook.save(f"tnledger[{date.replace('/','-')}].xlsx")

###################################################################################################

SERVICE_ACCOUNT_FILE = os.environ.get('VA_KEYS_FILE', 'keys.json')
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

creds = None
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

SAMPLE_SPREADSHEET_ID = '1DInbhKwN0FhePxkronq5mfp6gL6fmcSicn6Y4wE27Z0'


try:
    service = build('sheets', 'v4', credentials=creds)
    
    all_data = [list(t) for t in set(tuple(element) for element in all_data)]
    all_data.insert(0, [f'SELECTED DATE   =>   {selected_date}'])
    all_data.insert(0, [' - '])
    all_data.insert(2, [' - '])
                                                      
    sheet = service.spreadsheets()
    request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range="Sheet1!A2", valueInputOption="USER_ENTERED", body={"values":all_data}).execute()
                                                      
except HttpError as err:
    print(err)


driver.close()
print('Hey Angel Extraction is Completed! You can close this terminal.')


