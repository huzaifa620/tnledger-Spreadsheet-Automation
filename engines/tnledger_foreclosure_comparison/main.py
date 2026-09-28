# Generated from "tnledger foreclosure comparison.ipynb" so the VA hub can run it outside Jupyter.
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
import pandas as pd
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from openpyxl import Workbook
from openpyxl.styles import Font, Color, Alignment, Border, Side
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.utils import get_column_letter
import os, time, re, random
from openpyxl import load_workbook
import PySimpleGUI as sg
from datetime import date
from webdriver_manager.chrome import ChromeDriverManager

today = str(date.today()).split('-')
theme_name_list = sg.theme_list()

while True:
    sg.theme(theme_name_list[random.randint(0, len(theme_name_list)-1)])
    layout = [[
        sg.Text('Choose Default Date',
                size=(20, 1),
                font='Ubuntu',
                justification='left')
    ],
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
                  sg.Text('Choose Ending Date',
                          size=(20, 1),
                          font='Ubuntu',
                          justification='left')
              ],
              [
                  sg.Input(key='-IN5-', size=(20, 1)),
                  sg.CalendarButton(
                      'Calendar2',
                      font="Ubuntu",
                      target='-IN5-',
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
        if v['-IN5-'] == None or v['-IN5-'] == '' or v['-IN4-'] == None or v['-IN4-'] == '':
            print('Enter the date correctly')
            win.close()
            continue
        else:
            win.close()
            break
            
currentdate = time.strptime(v['-IN5-'].split(' ')[0], "%Y-%m-%d")
link = f'https://www.tnledger.com/Notices.aspx?noticesDate={currentdate.tm_mon}/{currentdate.tm_mday}/{currentdate.tm_year}'
print(link)
hardcodedate = time.strptime(v['-IN4-'].split(' ')[0], "%Y-%m-%d")
dates = link.split('Date=')[1]
dates = time.strptime(dates, "%m/%d/%Y")
dates_list = []
bold_font = Font(bold=True)
bold_font = Font(bold=True)
center_aligned_text = Alignment(horizontal="center")
double_border_side = Side(border_style="double")
square_border = Border(top=double_border_side, right=double_border_side, bottom=double_border_side, left=double_border_side)
workbook = Workbook()
sheet = workbook.active
sheet["A1"] = "Borrower's Name"
sheet["A1"].font = bold_font
sheet["A1"].alignment = center_aligned_text
sheet["A1"].border = square_border
sheet["B1"] = "Address"
sheet["B1"].font = bold_font
sheet["B1"].alignment = center_aligned_text
sheet["B1"].border = square_border
sheet["C1"] = "Remaining Address"
sheet["C1"].font = bold_font
sheet["C1"].alignment = center_aligned_text
sheet["C1"].border = square_border
sheet["D1"] = "Count"
sheet["D1"].font = bold_font
sheet["D1"].alignment = center_aligned_text
sheet["D1"].border = square_border
dim_holder = DimensionHolder(worksheet=sheet)

for col in range(sheet.min_column, sheet.max_column + 1):
    dim_holder[get_column_letter(col)] = ColumnDimension(sheet, min=col,max=col, width=20)

sheet.column_dimensions = dim_holder
row = 2
column = 1
files = [f for f in os.listdir('.') if os.path.isfile(f)]
for f in files:
    if 'xlsx' in f:
        if 'tnledger' in f:
            print(f)
            dates_list.append(
                re.findall('[[]\d+[-]\d+[-]\d+[]]',
                           f)[0].replace('[', '').replace(']', ''))

if link.split('Date=')[1].replace('/', '-') in dates_list:
    print("already exist")
else:
    options = webdriver.ChromeOptions()
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.headless = False
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-extensions")    
    options.add_argument('--disable-useAutomationExtension')
    cond = ""
    less_record = []
    for j in dates_list:
        da = time.strptime(j, "%m-%d-%Y")
        print(da)
        if da == dates:
            cond = "EQUAL"
        elif da < dates:
            cond = "GREATER"
            less_record.append(da)
        elif da > dates:
            cond = "LESS"

    if hardcodedate == dates:
        cond = "EQUAL"
    if len(less_record) == 0:
        if hardcodedate < dates or hardcodedate == dates:

            driver = webdriver.Chrome(executable_path=ChromeDriverManager().install(),chrome_options=options)
            driver.get(link)
            driver1 = webdriver.Chrome(executable_path=ChromeDriverManager().install(),chrome_options=options)
            driver.implicitly_wait(30)
            driver1.implicitly_wait(30)
            links_of_back = []
            links_of_back.append(link)
            while True:
                if cond == "EQUAL":
                    break
                l = driver.find_element(By.XPATH,
                    '//div[@id="left-column"]/div[4]/p/a').get_attribute(
                        'href')
                links_of_back.append(l)
                if time.strptime(l.split('Date=')[1], "%m/%d/%Y") == hardcodedate:
                    break

                driver.get(l)
            if len(links_of_back) > 1:
                reversed_list = links_of_back[::-1]
            else:
                reversed_list = links_of_back
            data = {}
            for u in reversed_list:
                print(u)
                dd = u.split('Date=')[1]
                driver.get(u)

                date = driver.current_url[driver.current_url.find('=') + 1:]
                all_data = []
                date = driver.current_url[driver.current_url.find('=') + 1:]
                forclosure_rows = driver.find_elements(By.XPATH,
                    '//table[@id="ContentPane_ForeclosureGridView"]/tbody/tr')
                del forclosure_rows[0]
                for i in forclosure_rows:
                    individual_rows = i.find_elements(By.TAG_NAME,'td')
                    each_row_data = []
                    for j in (individual_rows):
                        if j.text == "View":
                            current_link = (j.find_element(By.TAG_NAME,
                                'a').get_attribute('href'))
                            link_id = current_link[current_link.find('(') + 1:current_link.find(')')].split(',')[0].replace("'", "")
                            
                            driver1.get(f"https://www.tnledger.com/Search/Details/ViewNotice.aspx?id={link_id}&date={date}")
                            
                            table_data = driver1.find_elements(By.XPATH, '//table[@class="data-table"]/tbody/tr')
                            for counter, k in enumerate(table_data):
                                horizontal_data = k.find_elements(By.TAG_NAME,
                                    'td')
                                check = horizontal_data.pop(0).text
                                if 'Original Trustee:' in check:
                                    break
                                for l in horizontal_data:
                                    each_row_data.append(l.text)
                            all_data.append(each_row_data)

                data[dd] = all_data
            all_keys = []
            for key in data:
                all_keys.append(key)
            if len(data.keys()) == 1:
                for key in data:

                    for q in data[key]:
                        sheet.cell(row=row, column=column).value = q[0]
                        column += 1
                        sheet.cell(row=row, column=column).value = q[1]
                        column += 1
                        sheet.cell(row=row, column=column).value = q[2]
                        column += 1
                        sheet.cell(row=row, column=column).value = 0
                        column = 1
                        row += 1
            else:
                lis_of_res = []
                list_of_count = []
                a = 0
                while True:
                    c = 0
                    while c < len(data[all_keys[a]]):
                        if data[all_keys[a]][c] in lis_of_res:
                            index = lis_of_res.index(data[all_keys[a]][c])
                            count = list_of_count[index]
                            list_of_count[index] = count + 1
                        else:
                            lis_of_res.append(data[all_keys[a]][c])
                            list_of_count.append(0)
                        c += 1
                    a += 1
                    if a > len(all_keys) - 1:
                        break
                sold_list = []
                a = 0
                while True:
                    c = 0
                    while c < len(data[all_keys[a]]):
                        if data[all_keys[a]][c] not in data[all_keys[a + 1]]:
                            if data[all_keys[a]][c] in sold_list:
                                c += 1
                                continue
                            sold_list.append(data[all_keys[a]][c])
                        c += 1
                    a += 1
                    if a == len(all_keys) - 1:
                        break
                c = 0
                while c < len(sold_list):
                    if sold_list[c] in lis_of_res:
                        index = lis_of_res.index(sold_list[c])
                        lis_of_res.pop(index)
                        list_of_count.pop(index)
                    else:
                        print("ham")
                    c += 1
                for k, l in enumerate(lis_of_res):
                    sheet.cell(row=row, column=column).value = l[0]
                    column += 1
                    sheet.cell(row=row, column=column).value = l[1]
                    column += 1
                    sheet.cell(row=row, column=column).value = l[2]
                    column += 1
                    sheet.cell(row=row, column=column).value = list_of_count[k]
                    column = 1
                    row += 1
            driver.close()
            driver1.close()
            workbook.save(f"tnledger_foreclosure[{link.split('Date=')[1].replace('/','-')}].xlsx")
            print('Completed')
    else:
        greatest_one = less_record[0]
        for i in less_record:
            if i > greatest_one:
                greatest_one = i
        links_of_back = []
        links_of_back.append(link)

        driver = webdriver.Chrome(executable_path=ChromeDriverManager().install(),chrome_options=options)
        driver.get(link)
        driver1 = webdriver.Chrome(executable_path=ChromeDriverManager().install(),chrome_options=options)
        driver.implicitly_wait(30)
        driver1.implicitly_wait(30)
        if greatest_one == hardcodedate:
            given = 'EQUAL'
        elif greatest_one > hardcodedate:
            given = "GREATEST"
        elif greatest_one < hardcodedate:
            given = "HARDCODE"
        read_excel = False
        if given == 'EQUAL' or given == "GREATEST":
            read_excel = True
            while True:
                l = driver.find_element(By.XPATH,
                    '//div[@id="left-column"]/div[4]/p/a').get_attribute(
                        'href')

                if time.strptime(l.split('Date=')[1], "%m/%d/%Y") == greatest_one:
                    break
                driver.get(l)
        elif given == "HARDCODE":
            print("owias")
            while True:
                l = driver.find_element(By.XPATH,
                    '//div[@id="left-column"]/div[4]/p/a').get_attribute(
                        'href')
                links_of_back.append(l)
                if time.strptime(l.split('Date=')[1],
                                 "%m/%d/%Y") == hardcodedate:
                    break

                driver.get(l)

        if len(links_of_back) > 1:
            reversed_list = links_of_back[::-1]
        else:
            reversed_list = links_of_back
        data = {}
        if read_excel:
            name = f"tnledger_foreclosure[{greatest_one.tm_mon}-{greatest_one.tm_mday}-{greatest_one.tm_year}].xlsx"
            wb = load_workbook(filename=name)
            sheet = wb.active
            row_count = sheet.max_row
            column_count = sheet.max_column
            lst = []
            lst_count = []
            for i in range(2, row_count + 1):

                cond = False
                first = sheet.cell(row=i, column=1).value
                second = sheet.cell(row=i, column=2).value
                third = sheet.cell(row=i, column=3).value
                count = sheet.cell(row=i, column=4).value
                lst_count.append(count)
                lst.append([first, second, third])
            data[
                f'{greatest_one.tm_mon}/{greatest_one.tm_mday}/{greatest_one.tm_year}'] = lst
        for u in reversed_list:
            print(u)
            dd = u.split('Date=')[1]
            driver.get(u)

            date = driver.current_url[driver.current_url.find('=') + 1:]
            all_data = []
            date = driver.current_url[driver.current_url.find('=') + 1:]
            forclosure_rows = driver.find_elements(By.XPATH,
                '//table[@id="ContentPane_ForeclosureGridView"]/tbody/tr')
            del forclosure_rows[0]
            for i in forclosure_rows:
                individual_rows = i.find_elements(By.TAG_NAME,'td')
                each_row_data = []
                for j in (individual_rows):
                    if j.text == "View":
                        current_link = (j.find_element(By.TAG_NAME,
                            'a').get_attribute('href'))
                        link_id = current_link[current_link.find('(') +
                                               1:current_link.find(')')].split(
                                                   ',')[0].replace("'", "")

                        driver1.get(f"https://www.tnledger.com/Search/Details/ViewNotice.aspx?id={link_id}&date={date}")
                        table_data = driver1.find_elements(By.XPATH,
                            '//table[@class="data-table"]/tbody/tr')
                        for counter, k in enumerate(table_data):
                            horizontal_data = k.find_elements(By.TAG_NAME,'td')
                            check = horizontal_data.pop(0).text
                            if 'Original Trustee:' in check:
                                break
                            for l in horizontal_data:
                                each_row_data.append(l.text)
                all_data.append(each_row_data)

            data[dd] = all_data
        all_keys = []
        for key in data:
            all_keys.append(key)
        if len(data.keys()) == 1:
            for key in data:
                for q in data[key]:
                    sheet.cell(row=row, column=column).value = q[0]
                    column += 1
                    sheet.cell(row=row, column=column).value = q[1]
                    column += 1
                    sheet.cell(row=row, column=column).value = q[2]
                    column += 1
                    sheet.cell(row=row, column=column).value = 0
                    column = 1
                    row += 1
        else:
            lis_of_res = []
            list_of_count = []
            a = 0
            while True:
                c = 0
                while c < len(data[all_keys[a]]):
                    if a == 0 and read_excel == True:
                        if data[all_keys[a]][c] in lis_of_res:
                            index = lis_of_res.index(data[all_keys[a]][c])
                            count = list_of_count[index]
                            list_of_count[index] = count + 1
                        else:
                            lis_of_res.append(data[all_keys[a]][c])
                            list_of_count.append(lst_count[c])
                    else:

                        if data[all_keys[a]][c] in lis_of_res:
                            index = lis_of_res.index(data[all_keys[a]][c])
                            count = list_of_count[index]
                            list_of_count[index] = count + 1
                        else:
                            lis_of_res.append(data[all_keys[a]][c])
                            list_of_count.append(0)
                    c += 1
                a += 1
                if a > len(all_keys) - 1:
                    break
            sold_list = []
            a = 0
            while True:
                c = 0
                while c < len(data[all_keys[a]]):
                    if data[all_keys[a]][c] not in data[all_keys[a + 1]]:
                        if data[all_keys[a]][c] in sold_list:
                            c += 1
                            continue
                        sold_list.append(data[all_keys[a]][c])
                    c += 1
                a += 1
                if a == len(all_keys) - 1:
                    break
            c = 0
            while c < len(sold_list):
                if sold_list[c] in lis_of_res:
                    index = lis_of_res.index(sold_list[c])
                    lis_of_res.pop(index)
                    list_of_count.pop(index)
                else:
                    print("ham")
                c += 1
            for k, l in enumerate(lis_of_res):
                sheet.cell(row=row, column=column).value = l[0]
                column += 1
                sheet.cell(row=row, column=column).value = l[1]
                column += 1
                sheet.cell(row=row, column=column).value = l[2]
                column += 1
                sheet.cell(row=row, column=column).value = list_of_count[k]
                column = 1
                row += 1
        driver.close()
        driver1.close()
        workbook.save(f"tnledger_foreclosure[{link.split('Date=')[1].replace('/','-')}].xlsx")
        pd.DataFrame(sold_list,columns=['Name','Address','Remaining Address']).to_excel("unrepeated_data.xlsx", index = False, header=True)
        print('Completed')

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2 import service_account

SERVICE_ACCOUNT_FILE = os.environ.get('VA_KEYS_FILE', 'keys.json')
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

creds = None
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

SAMPLE_SPREADSHEET_ID = '1k0j25pNfx6hjKmANxjEPLkLduWuSa3-csb8pTU17TvQ'

try:
    service = build('sheets', 'v4', credentials=creds)
                                                      
    sheet = service.spreadsheets()
    
    request = sheet.values().insert_rows(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                        range="Sheet1!A1", valueInputOption="USER_ENTERED", body={"values":data}).execute()
    
except HttpError as err:
    print(err)
