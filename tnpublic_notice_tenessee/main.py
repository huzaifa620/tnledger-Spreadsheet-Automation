from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
import undetected_chromedriver as uc
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from openpyxl import Workbook
from openpyxl.styles import Font, Color, Alignment, Border, Side
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.utils import get_column_letter
from tkinter import simpledialog
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from datetime import date
import tkinter as tk
from webdriver_manager.chrome import ChromeDriverManager
import PySimpleGUI as sg
from byerecaptcha import solveRecaptcha
import nltk, urllib, random, requests, csv, os, sys, time, re, datefinder
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2 import service_account

SERVICE_ACCOUNT_FILE = 'keys.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

creds = None
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

SAMPLE_SPREADSHEET_ID = '1278HJam-TruXlKrlHA3mdTHtUEljeK-xEC4co0kytaA'

today = date.today()
theme_name_list = sg.theme_list()
today = str(date.today()).split('-')

court_names = {'mon':'Montgomery','dav':'Davidson','rob':'Robertson','wil':'Wilson','rut':'Rutherford'}
while True:
    sg.theme(theme_name_list[random.randint(0, len(theme_name_list)-1)])
    #define layout
    layout=[
            [sg.Text('Enter the Starting date',size=(20, 1), font='Ubuntu',justification='left')],
            [sg.Input(key='from', size=(20,1)), sg.CalendarButton('Calendar',font="Ubuntu",  target='from', default_date_m_d_y=(int(today[1]),int(today[2]),int(today[0])), )],
            [sg.Text('Enter the Ending date',size=(20, 1), font='Ubuntu',justification='left')],
            [sg.Input(key='to', size=(20,1)), sg.CalendarButton('Calendar',font="Ubuntu",  target='to', default_date_m_d_y=(int(today[1]),int(today[2]),int(today[0])), )],
            [sg.Frame(' Select County ',[[sg.Radio('Montgomery', default=True, key="mon",group_id='2', font = 'Ubuntu')],[sg.Radio('Davidson', default=True, key="dav",group_id='2', font = 'Ubuntu')],[sg.Radio('Robertson', default=False, key="rob",group_id='2', font = 'Ubuntu')],[sg.Radio('Rutherford', default=False, key="rut",group_id='2', font = 'Ubuntu')],[sg.Radio('Wilson', default=False, key="wil",group_id='2', font = 'Ubuntu')]],border_width=3,font = 'Ubuntu',relief = "solid")],
            [sg.Button('OK', font=('Ubuntu',12)),sg.Button('CANCEL', font=('Ubuntu',12))]]
    #Define Window
    win =sg.Window('TnPublic Notices Foreclosures',layout)
    #Read  values entered by user
    e,v=win.read()
    con = False 
    if e == None or e == "CANCEL":
        starting_date_entry=None
        ending_date_entry=None
        county=None
        win.close()
        con = True

        break
    else:
        if  v['from'] == None or v['from'] == '' or v['to'] == None or v['to'] == '' :
            print('Enter the date correctly')
            win.close()
            continue
        elif v['mon'] == False and v['rut'] == False and v['rob'] == False and  v['wil'] == False and  v['dav'] == False:
            print('please select the radio button')
            win.close()
            continue
        else:
            starting_date_entry = f"{v['from'].split(' ')[0].split('-')[1]}/{v['from'].split(' ')[0].split('-')[2]}/{v['from'].split(' ')[0].split('-')[0]}"
            ending_date_entry = f"{v['to'].split(' ')[0].split('-')[1]}/{v['to'].split(' ')[0].split('-')[2]}/{v['to'].split(' ')[0].split('-')[0]}"
            for key in v:
                if key == 'mon' or key=='rut' or key == 'rob' or key=='wil' or key=='dav':
                    if v[key] == True:
                        county=court_names[key]
                        break
            win.close()
            break

            
if starting_date_entry==None or ending_date_entry==None or county==None:
    raise Exception('Please select the starting and ending date with a county name to start the scrapper.')
    
print(starting_date_entry)
print(ending_date_entry)
print(county)



options = webdriver.ChromeOptions() 
driver = uc.Chrome(executable_path=ChromeDriverManager().install(),use_subprocess=True,options=options)
driver.get('https://www.tnpublicnotice.com/')
driver.maximize_window()
driver.implicitly_wait(30)
search_list=Select(driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_as1_ddlPopularSearches'))
search_list.select_by_visible_text('Probate Notices')
time.sleep(5)
driver.find_element(By.ID,"ctl00_ContentPlaceHolder1_as1_divCounty").click()
for i in driver.find_elements(By.XPATH,'//ul[@id="ctl00_ContentPlaceHolder1_as1_lstCounty"]/li'):
    if i.text.strip()==county.strip():
        ActionChains(driver).move_to_element(driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_as1_btnGo')).perform()
        print(i.text)
        i.click()
        ActionChains(driver).move_to_element(driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_as1_btnGo')).perform()
        break
time.sleep(2)
ActionChains(driver).move_to_element(driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_as1_btnGo')).perform()
driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_as1_divDateRange').click()
driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_as1_rbRange').click()
ActionChains(driver).move_to_element(driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_as1_txtDateFrom')).perform()
driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_as1_txtDateFrom').clear()
driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_as1_txtDateFrom').send_keys(starting_date_entry)
driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_as1_txtDateTo').clear()
driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_as1_txtDateTo').send_keys(ending_date_entry)
ActionChains(driver).move_to_element(driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_as1_btnGo')).perform()
driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_as1_btnGo').click()
time.sleep(5)
try:
    pages_viewable=Select(driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_WSExtendedGridNP1_GridView1_ctl01_ddlPerPage'))
    pages_viewable.select_by_visible_text('50')
except:
    pass
time.sleep(2)

all_data=[]
try:
    total_pages=int(driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_WSExtendedGridNP1_GridView1_ctl01_lblTotalPages').text.split()[1])
except:
    total_pages=1
for j in range(total_pages):
    for i in range(len(driver.find_elements(By.XPATH,"//td[@class='view bdrBrownTop bdrBrownRight bdrBrownBottom bdrBrownLeft mobileL mobileT']/input[1]"))):
        try:
            driver.find_elements(By.XPATH,f"//td[@class='view bdrBrownTop bdrBrownRight bdrBrownBottom bdrBrownLeft mobileL mobileT']/input[1]")[i].click()
        except IndexError as error:
            print('Done extraction')
            break
        while True:  
            try:
                if 'reCAPTCHA' in (driver.find_element(By.ID,"aspnetForm").text):
                    solveRecaptcha(driver)
                    driver.execute_script('''
                    try{
                        document.getElementById('ctl00_ContentPlaceHolder1_PublicNoticeDetailsBody1_btnViewNotice').click()}
                    catch{}
                        ''')
                    print(driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_PublicNoticeDetailsBody1_pnlNoticeContent').text)
                    all_data.append(driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_PublicNoticeDetailsBody1_pnlNoticeContent').text)
                    driver.find_element(By.XPATH,'//p[@class="backlink"]/a').click()
                    driver.implicitly_wait(30)
                    break
                else:
                    print(driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_PublicNoticeDetailsBody1_pnlNoticeContent').text)
                    all_data.append(driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_PublicNoticeDetailsBody1_pnlNoticeContent').text)
                    driver.back()
                    driver.implicitly_wait(30)
                    break
            except:
                try:
                    solveRecaptcha(driver)
                    driver.execute_script('''
                    try{
                        document.getElementById('ctl00_ContentPlaceHolder1_PublicNoticeDetailsBody1_btnViewNotice').click()}
                    catch{}
                        ''')
                    print(driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_PublicNoticeDetailsBody1_pnlNoticeContent').text)
                    all_data.append(driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_PublicNoticeDetailsBody1_pnlNoticeContent').text)
                    driver.find_element(By.XPATH,'//p[@class="backlink"]/a').click()
                    driver.implicitly_wait(30)
                    break
                except:
                    driver.switch_to.default_content()
                    ROOT = tk.Tk()
                    ROOT.withdraw()
                    driver.execute_script('''
                    try{
                        document.getElementById('ctl00_ContentPlaceHolder1_PublicNoticeDetailsBody1_btnViewNotice').click()
                        }
                    catch{}
                        ''')
                    try:    
                        print(driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_PublicNoticeDetailsBody1_pnlNoticeContent').text)
                    except:
                        USER_INP=simpledialog.askstring(title="Solve captcha manually", prompt="Have you done? Also please click the View notice button")
                    print(driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_PublicNoticeDetailsBody1_pnlNoticeContent').text)
                    all_data.append(driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_PublicNoticeDetailsBody1_pnlNoticeContent').text)
                    driver.find_element(By.XPATH,'//p[@class="backlink"]/a').click()
                    driver.implicitly_wait(30)
                    break
    try:
        if total_pages>1:    
            driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_WSExtendedGridNP1_GridView1_ctl54_btnNext').click()
            driver.implicitly_wait(30)
    except:
        break
        
if county.strip()=='Montgomery':
    notice_of_trustee_foreclosure_sale=[]
    substitue_trustee_sale=[]
    notice_of_trustee_sale=[]
    notice_to_creditors=[]
    notice_of_substitute_trustee_sale=[]
    notice_of_successor_trustee_sale=[]
    notice_of_foreclosure_sale=[]
    notice_of_substitute_trustee_foreclosure_sale=[]
    unfiltered_data=[]
    
    for i in all_data:
        if "NOTICE OF TRUSTEE'S FORECLOSURE SALE" in i:
            executor=None
            interested_parties=None
            property_address=None
            extracted_text=(nltk.tokenize.sent_tokenize(i))
            for j in extracted_text:
                if 'executed by' in j and executor==None:
                    executor=((j[j.index('executed by'):])[:(j[j.index('executed by'):]).index('for')].replace('executed by','').strip())
                if 'other interested parties' in j.lower() and interested_parties==None:
                    interested_parties=((j[j.index('Other Interested Parties'):])[:(j[j.index('Other Interested Parties'):]).index('The hereinafter')].replace('Other Interested Parties','').replace(':','').strip())
                if 'street address' in j.lower() and 'property is believed to be' in j.lower() and property_address==None:
                    property_address=((j[j.index('property is believed to be'):])[:(j[j.index('property is believed to be'):]).index(', but')].replace('property is believed to be','').replace(':','').strip())
                if executor!=None and interested_parties!=None and property_address!=None:
                    break
            notice_of_trustee_foreclosure_sale.append([executor,property_address.replace(':', ''),interested_parties])
        elif "SUBSTITUTE TRUSTEE'S SALE" in i:
            executor=None
            interested_parties=None
            property_address=None
            extracted_text=(nltk.tokenize.sent_tokenize(i))
            for j in extracted_text:
                if 'executed by' in j and executor==None:
                    executor=((j[j.index('executed by'):])[:(j[j.index('executed by'):]).index(',')].replace('executed by','').strip()).split('conveying')[0].split(' to ')[0].title().replace(';',' and')
                if 'enforce the debt' in j.lower() and interested_parties==None:         
                    interested_parties=((j[j.index('Enforce the Debt'):])[:-1].replace('Enforce the Debt','').replace(':','').strip())
                if 'OTHER INTERESTED PARTIES' in j.upper() and interested_parties==None:
                    interested_parties=''
                    keyword=False
                    for k in j.splitlines():
                        if 'OTHER INTERESTED PARTIES' in k.upper():
                            keyword=True
                        if 'The sale of the above-described property' in k or 'THIS IS AN ATTEMPT TO COLLECT A DEBT' in k.upper() or 'All right of equity of redemption' in k:
                            break
                        if keyword:
                            interested_parties=interested_parties+k
                    interested_parties=interested_parties.lower().replace('other interested parties','').replace(':','').strip().title()
                if 'interested parties may include' in j.lower() and interested_parties==None:
                    interested_parties=((j[j.index('interested parties may include'):])[:-1].replace('interested parties may include','').replace(':','').strip().replace(',','|'))
                if 'street address' in j.lower() and property_address==None:
                    if 'is believed to be' in j.lower():
                        property_address=((j[j.index('is believed to be'):])[:-1].replace('is believed to be','').replace(':','').strip())
                        if ' but such address is' in property_address:
                            property_address=property_address[:property_address.index(' but such address is')].replace(' but such address is','').replace(',','').strip()
                    if '?street address' in j.lower():
                        property_address=j.split(':')[-1].replace('?','').strip()
            substitue_trustee_sale.append([executor,property_address,interested_parties])
        elif "NOTICE OF TRUSTEE'S SALE" in i:
            executor=None
            interested_parties=None
            property_address=None
            extracted_text=(nltk.tokenize.sent_tokenize(i))
            for j in extracted_text:
                if 'executed by' in j and executor==None:
                    executor=((j[j.index('executed by'):])[:(j[j.index('executed by'):]).index(',')].replace('executed by','').strip()).split('conveying')[0].split(' to ')[0].title().replace(';',' and')
                if 'ALSO KNOWN AS' in j.upper() and property_address==None:
                    property_address=((j[j.index('ALSO KNOWN AS'):])[:-1].replace('ALSO KNOWN AS','').strip().splitlines()[0])
                if (('The sale held pursuant to this Notice' in j) or 'referenced property' in j) and interested_parties==None:
                    interested_parties=''
                    for k in (j.splitlines()):
                        if 'referenced property' in k or 'The sale held pursuant' in k:
                            pass
                        else:
                            interested_parties=interested_parties+k+' | '
                    if interested_parties.strip()[-1]=='|':
                        interested_parties=interested_parties.strip()[:-1].title()
            notice_of_trustee_sale.append([executor,property_address.replace(':', ''),interested_parties])
        elif 'NOTICE TO CREDITORS' in i:
            executor=None
            estate_of=None
            property_address=None
            attorney=None
            extracted_text=i.splitlines()
            for j in extracted_text:
                if 'estate of' in j.lower() and estate_of==None:
                    estate_of=j.split('(')[0].title().replace('Estate Of','').replace('?','').strip()
                if ('administrator' in j.lower() or 'personal representative' in j.lower() or 'executrix' in j.lower() or 'executor' in j.lower() or 'administratrix' in j.lower()):
                    if executor==None:
                        executor=j.split('-')[0].strip()
                    else:
                        executor=executor+' | '+j.split('-')[0].strip()
                if 'attorney' in j.lower() and attorney==None:
                    attorney=j.split(':')[-1].strip()
                if ',TN' in j or ', TN' in j and property_address==None:
                    property_address=''
                    factor=0
                    while True:
                        property_address=extracted_text[extracted_text.index(j)+factor]+' '+property_address
                        factor=factor-1
                        if ('attorney' or 'administrator'or 'personal representative'or 'executrix' or 'executor'or 'administratrix') in extracted_text[extracted_text.index(j)+factor].lower():
                            break
            notice_to_creditors.append([estate_of,executor,property_address,attorney])
        elif 'NOTICE OF SUBSTITUTE TRUSTEES SALE' in i:
            executor=None
            property_address=None
            extracted_text=i.splitlines()
            for j in extracted_text:
                if 'owner of property' in j.lower() and executor==None:
                    executor=(j.split(':')[-1].split(',')[0])
                if 'believed to be' in j.lower() and 'street address' in j.lower() and property_address==None:
                    property_address=((j[j.index('property is believed to be'):])[:(j[j.index('property is believed to be'):]).index('but')].replace('property is believed to be','').replace(':','').strip())
            notice_of_substitute_trustee_sale.append([executor,property_address])
        elif "NOTICE OF SUCCESSOR TRUSTEE'S SALE" in i:
            executor=None
            property_address=None
            interested_parties=None
            extracted_text=(nltk.tokenize.sent_tokenize(i))
            for j in extracted_text:
                if 'executed by' in j.lower() and executor==None:
                    executor=(((j[j.index('executed by'):])[:(j[j.index('executed by'):]).index(',')].replace('executed by','').strip())).title()
                if 'commonly known as' in j.lower() and property_address==None:
                    property_address=(j.split(':')[-1]).strip()
                if 'claim' in j.lower() and interested_parties==None:
                    interested_parties=(((j[j.index('referenced property'):])[:-1].replace('referenced property','').strip()))
            notice_of_successor_trustee_sale.append([executor,property_address,interested_parties])
        elif 'NOTICE OF FORECLOSURE SALE OF REAL ESTATE' in i:
            executor=None
            property_address=None
            interested_parties=None
            attorney=None
            extracted_text=i.splitlines()
            for j in extracted_text:
                if 'deed of trust' in j.lower() and executor==None:
                    executor=(((j[j.index('"Deed of Trust"),'):])[:(j[j.index('"Deed of Trust"),'):]).index('and')].replace('"Deed of Trust"),','').strip())).title()
                if 'interested parties are' in j.lower() and interested_parties==None:
                    interested_parties=j.split(':')[-1].strip()
                if 'property address' in j.lower() and 'subject of this notice is' in j.lower() and property_address==None:
                    property_address=j.split(':')[-1].strip()
                if 'attorney' in j.lower() and attorney==None:
                    attorney=(extracted_text[extracted_text.index(j)-1])
            notice_of_foreclosure_sale.append([executor,interested_parties,property_address,attorney])

        elif 'NOTICE OF SUBSTITUTE TRUSTEE`S FORECLOSURE SALE' in i:
            extracted_text=i.splitlines()
            executor=None
            property_address=None
            interested_parties=None
            for j in extracted_text:
                if 'street address' in j.lower() and property_address==None:
                    if 'is believed to be' in j.lower():
                        property_address=((j[j.index('is believed to be'):])[:-1].replace('is believed to be','').replace(':','').strip())
                        if ' but such address is' in property_address:
                            property_address=property_address[:property_address.index(' but such address is')].replace(' but such address is','').replace(',','').strip()
                    if '?street address' in j.lower():
                        property_address=j.split(':')[-1].replace('?','').strip()
                if 'owner' in j.lower() and executor==None:
                    executor=(j.split(':')[-1])
                if 'interested part' in j.lower() and interested_parties==None:
                    interested_parties=(j.split(':')[-1])
            notice_of_substitute_trustee_foreclosure_sale.append([executor,property_address,interested_parties])
        else:
            unfiltered_data.append([i])
            

    notice_of_trustee_foreclosure_sale = [list(t) for t in set(tuple(element) for element in notice_of_trustee_foreclosure_sale)]
    notice_of_trustee_foreclosure_sale.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    notice_of_trustee_foreclosure_sale.insert(0, [' - '])
    notice_of_trustee_foreclosure_sale.insert(2, [' - '])
    
    substitue_trustee_sale = [list(t) for t in set(tuple(element) for element in substitue_trustee_sale)]
    substitue_trustee_sale.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    substitue_trustee_sale.insert(0, [' - '])
    substitue_trustee_sale.insert(2, [' - '])
    
    notice_of_trustee_sale = [list(t) for t in set(tuple(element) for element in notice_of_trustee_sale)]
    notice_of_trustee_sale.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    notice_of_trustee_sale.insert(0, [' - '])
    notice_of_trustee_sale.insert(2, [' - '])
    
    notice_to_creditors = [list(t) for t in set(tuple(element) for element in notice_to_creditors)]
    notice_to_creditors.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    notice_to_creditors.insert(0, [' - '])
    notice_to_creditors.insert(2, [' - '])
    
    notice_of_substitute_trustee_sale = [list(t) for t in set(tuple(element) for element in notice_of_substitute_trustee_sale)]
    notice_of_substitute_trustee_sale.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    notice_of_substitute_trustee_sale.insert(0, [' - '])
    notice_of_substitute_trustee_sale.insert(2, [' - '])
    
    notice_of_successor_trustee_sale = [list(t) for t in set(tuple(element) for element in notice_of_successor_trustee_sale)]
    notice_of_successor_trustee_sale.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    notice_of_successor_trustee_sale.insert(0, [' - '])
    notice_of_successor_trustee_sale.insert(2, [' - '])
    
    notice_of_foreclosure_sale = [list(t) for t in set(tuple(element) for element in notice_of_foreclosure_sale)]
    notice_of_foreclosure_sale.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    notice_of_foreclosure_sale.insert(0, [' - '])
    notice_of_foreclosure_sale.insert(2, [' - '])
    
    notice_of_substitute_trustee_foreclosure_sale = [list(t) for t in set(tuple(element) for element in notice_of_substitute_trustee_foreclosure_sale)]
    notice_of_substitute_trustee_foreclosure_sale.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    notice_of_substitute_trustee_foreclosure_sale.insert(0, [' - '])
    notice_of_substitute_trustee_foreclosure_sale.insert(2, [' - '])
    
    unfiltered_data = [list(t) for t in set(tuple(element) for element in unfiltered_data)]
    unfiltered_data.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    unfiltered_data.insert(0, [' - '])
    unfiltered_data.insert(2, [' - '])
    
    try:
        service = build('sheets', 'v4', credentials=creds)

        sheet = service.spreadsheets()
        

        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range="Montgomery Notice of Trustee Foreclosure Sale!A2", valueInputOption="USER_ENTERED", body={"values":notice_of_trustee_foreclosure_sale}).execute()

        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range="Montgomery Substitue Trustee Sale!A2", valueInputOption="USER_ENTERED", body={"values":substitue_trustee_sale}).execute()
        
        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range="Montgomery Notice of Trustee Sale!A2", valueInputOption="USER_ENTERED", body={"values":notice_of_trustee_sale}).execute()
        
        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range="Montgomery Notice to Creditors!A2", valueInputOption="USER_ENTERED", body={"values":notice_to_creditors}).execute()
        
        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range="Montgomery Notice of Substitute Trustee Sale!A2", valueInputOption="USER_ENTERED", body={"values":notice_of_substitute_trustee_sale}).execute()
        
        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range="Montgomery Notice of Successor Trustee Sale!A2", valueInputOption="USER_ENTERED", body={"values":notice_of_successor_trustee_sale}).execute()
        
        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range="Montgomery Notice of Foreclosure Sale!A2", valueInputOption="USER_ENTERED", body={"values":notice_of_foreclosure_sale}).execute()
        
        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range="Montgomery Notice of Substitute Trustee Foreclosure Sale!A2", valueInputOption="USER_ENTERED", body={"values":notice_of_substitute_trustee_foreclosure_sale}).execute()
        
        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range="Montgomery Unfiltered Data!A2", valueInputOption="USER_ENTERED", body={"values":unfiltered_data}).execute()

    except HttpError as err:
        print(err)

if county.strip()=='Davidson':
    notice_to_creditors=[]
    notice_of_successor_trustee_sale=[]
    substitute_trustee_sale=[]
    notice_of_trustee_sale=[]
    trustee_sale=[]
    substitute_trustee_notice_of_sale=[]
    notice_of_substitute_trustee_sale=[]
    substitute_trustee_sale=[]
    notice_of_foreclosure_sale_state=[]
    unfiltered_data=[]

    def containsNumberorLetterNumber(input):
        has_letter = False
        has_number = False
        for x in input:
            if x.isalpha():
                has_letter = True
            elif x.isnumeric():
                has_number = True
            if has_number:
                return True
        return False

    for i in all_data:
        if 'NOTICE TO CREDITORS' in i and 'Docket No' in list(filter(None, i.splitlines()))[2]:
            extracted_text=i.splitlines()
            executor=None
            executor_address=None
            executor_with_address=None
            estate_of=None
            attorney=None
            for j in extracted_text:
                if 'estate of' in j.lower() and estate_of==None:
                    estate_of=(j.split(',')[0].replace('Estate of','').strip())
                if 'personal representative' in j.lower() and attorney==None:
                    attorney=extracted_text[extracted_text.index(j)+1]

                if ('day of' and 'This')  in j and executor_with_address==None:
                    executor_with_address=''
                    factor=1
                    while True:
                        executor_with_address=executor_with_address+' '+(extracted_text[extracted_text.index(j)+factor])
                        if len(re.findall(', [A-Z]{2,3} \d{5}',(extracted_text[extracted_text.index(j)+factor]))):
                            executor_with_address=executor_with_address+' | '
                        factor=factor+1
                        if 'personal representative' in extracted_text[extracted_text.index(j)+factor].lower():
                            break

                    executor=''
                    executor_address=''
                    for l in executor_with_address.strip().split('|'):
                        for k in l.split():
                            if 'P.O.' not in k and k.replace('.','').isalpha() :
                                executor=executor+k+' '
                            else:
                                executor=executor+' | '
                                break
                    if executor.strip()[-1]=="|":
                        executor=executor.strip()[:-1].title()

                    for l in executor_with_address.strip().split('|'):
                        address_found=False
                        for k in l.split():
                            if '|' in k:
                                break
                            elif address_found:
                                executor_address=executor_address+k+' '
                            elif containsNumberorLetterNumber(k):
                                address_found=True
                                executor_address=executor_address+k+' '
                        executor_address=executor_address+' | '
                    executor_address=executor_address.replace('|  |','')

            notice_to_creditors.append([estate_of,executor,executor_address,attorney])
        elif "NOTICE OF SUCCESSOR TRUSTEE'S SALE" in i:
            extracted_text=i.splitlines()
            executor=None
            interested_parties=None
            property_address=None
            for j in extracted_text:

                if 'executed by' in j and executor==None:
                    executor=((j[j.index('executed by'):])[:(j[j.index('executed by'):]).index(',')].replace('executed by','').strip())   
                    if 'and' in executor:
                        executor=executor.lower().replace('and','|').title()

                if 'current owner' in j.lower() and executor==None:
                    executor=(j.splitlines()[0].split(':')[-1])

                if 'interested part' in j.lower() and interested_parties==None:
                    if 'Publish' in j:
                        interested_parties=((j[j.index('INTERESTED PARTIES: '):])[:(j[j.index('INTERESTED PARTIES: '):]).index('Publish')].replace('INTERESTED PARTIES: ','').strip())   
                    else:
                        interested_parties=(j.split(':')[-1])

                if 'address' in j.lower() and property_address==None:
                    if 'believed to be' in j.lower():
                        property_address=(((j[j.index('believed to be'):])[:(j[j.index('believed to be'):]).index(', but')].replace('believed to be','').strip())   )
                    if 'Address' in j and 'Reference' in j:
                        property_address=(((j[j.index('Address'):])[:(j[j.index('Address'):]).index('Reference')].replace('Address','').strip()))

            notice_of_successor_trustee_sale.append([executor,property_address,interested_parties])
        elif "SUBSTITUTE TRUSTEE'S SALE" in i:
            extracted_text=i.splitlines()
            executor=None
            interested_parties=None
            property_address=None
            for j in extracted_text:
                if 'executed by' in j.lower() or 'CURRENT PROPERTY OWNER:' in j and executor==None:
                    if 'executed by' in j.lower():
                        executor=((j[j.index('executed by'):])[:(j[j.index('executed by'):]).index(',')].replace('executed by','').strip())
                    if 'CURRENT PROPERTY OWNER:' in j:
                        executor=j.split(':')[-1].strip()

                if 'address' in j.lower() or 'commonly known as' in j.lower() and property_address==None:
                    property_address=j.split(':')[-1].strip()

                if 'Such parties known to the Substitute Trustee may include' in j or 'entities have an interest in the above-described property' in j or 'OTHER LIEN HOLDERS OR HOLDERS OF INTEREST' in j and interested_parties==None:
                    if 'Such parties known to the Substitute Trustee may include' in j:
                        interested_parties=((j[j.index('Such parties known to the Substitute Trustee may include'):])[:-1].replace('Such parties known to the Substitute Trustee may include','').replace(':','').strip())

                    elif 'OTHER LIEN HOLDERS OR HOLDERS OF INTEREST' in j:
                        interested_parties=(j.split(':')[-1])

                    else:
                        interested_parties=(j.split(':')[-1])

            substitute_trustee_sale.append([executor,property_address,interested_parties])
        elif "NOTICE OF TRUSTEE'S SALE" in i:
                extracted_text=i.splitlines()
                executor=None
                property_address=None
                interested_parties=None
                no_of_parties=None
                for j in extracted_text:
                    if 'executed by' in j and executor==None:
                        executor=((j[j.index('executed by'):])[:(j[j.index('executed by'):]).index(',')].replace('executed by','').strip())
                    if 'ALSO KNOWN AS' in j and property_address==None:
                        property_address=(j.replace('ALSO KNOWN AS:','').strip())
                    if ('The sale held pursuant to this Notice' in j) and interested_parties==None:
                        count=1
                        factor=1
                        if 'On or about' in extracted_text[extracted_text.index(j)-1]:
                            count=2
                            factor=2
                        interested_parties=''
                        while True:
                            interested_parties+=(extracted_text[extracted_text.index(j)-count])+' | '
                            count=count+1
                            if 'referenced property:' in extracted_text[extracted_text.index(j)-count]:
                                no_of_parties=count-factor
                                break
                if executor==None or executor=='None':
                    unfiltered_data.append([i])
                else:
                    notice_of_trustee_sale.append([executor,property_address,interested_parties,no_of_parties])

        elif "TRUSTEE'S SALE" in i:
            extracted_text=i.splitlines()
            executor=None
            property_address=None
            interested_parties=None
            for j in extracted_text:
                if 'executed by' in j and executor==None:
                    executor=((j[j.index('executed by'):])[:(j[j.index('executed by'):]).index(',')].replace('executed by','').strip())
                if 'address' in j.lower() and property_address==None:
                    property_address=((j[j.index(':'):])[:(j[j.index(':'):]).index('as shown on')].replace(':','').strip())
                if 'interested part' in j.lower() and interested_parties==None:
                    interested_parties=(j.split(':')[-1])
            if executor==None or executor=='None':
                unfiltered_data.append([i])
            else:
                trustee_sale.append([executor,property_address,interested_parties])

        elif "SUBSTITUTE TRUSTEE'S NOTICE OF SALE" in i:
            extracted_text=i.splitlines()
            executor=None
            property_address=None
            interested_parties=None
            for j in extracted_text:
                if 'executed by' in j and executor==None:
                    executor=((j[j.index('executed by'):])[:(j[j.index('executed by'):]).index(',')].replace('executed by','').strip())
                if 'address' in j.lower() and property_address==None:
                    if 'believed to be' in j.lower():
                        property_address=(((j[j.index('believed to be'):])[:(j[j.index('believed to be'):]).index(', but')].replace('believed to be','').strip())   )
                if 'interested part' in j.lower() and interested_parties==None:
                    interested_parties=j.split(':')[-1]
            if executor==None or executor=='None':
                unfiltered_data.append([i])
            else:
                substitute_trustee_notice_of_sale.append([executor,property_address,interested_parties])

        elif "NOTICE OF SUBSTITUTE TRUSTEE S SALE" in i:
            extracted_text=(nltk.tokenize.sent_tokenize(i))
            executor=None
            property_address=None
            interested_parties=None
            for j in extracted_text:
                if 'executed by' in j and executor==None:
                    executor=((j[j.index('executed by'):])[:(j[j.index('executed by'):]).index('conveying')].replace('executed by','').strip())
                if 'street address' in j.lower() and property_address==None:
                    if 'believed to be' in j.lower():
                        property_address=(((j[j.index('believed to be'):])[:-1].replace('believed to be','').strip())   ) 
                if 'interested part' in j.lower() and interested_parties==None:
                    interested_parties=((j[j.index('PARTIES:'):])[:(j[j.index('PARTIES:'):]).index('The sale')].replace('PARTIES:','').strip())

            notice_of_substitute_trustee_sale.append([executor,property_address,interested_parties])
        elif "SUBSTITUTE TRUSTEES SALE" in  i:
            extracted_text_1=i.splitlines()
            extracted_text_2=(nltk.tokenize.sent_tokenize(i))
            executor=None
            property_address=None
            interested_parties=None
            for j in extracted_text_2:
                if 'street address' in j.lower() and property_address==None:
                    if 'is believed to be' in j.lower():
                        property_address=((j[j.index('is believed to be'):])[:(j[j.index('is believed to be'):]).index(', but such address')].replace('is believed to be','').strip())
                    elif 'commonly known as' in j.lower():
                        property_address=((j[j.index('Commonly known as'):])[:(j[j.index('Commonly known as'):]).index('Parcel')].replace('Commonly known as','').replace(':', '').strip())
                    else:
                        property_address=((j[j.index('Street Address:'):])[:(j[j.index('Street Address:'):]).index('Parcel')].replace('Street Address:','').strip())               
                if 'INTERESTED PARTIES' in j and interested_parties==None:
                    if 'THIS IS AN ATTEMPT' in j:
                        interested_parties=((j[j.index('PARTIES:'):])[:(j[j.index('PARTIES:'):]).index('THIS IS AN ATTEMPT')].replace('PARTIES:','').strip())               
                    else:
                        interested_parties=((j[j.index('PARTIES:'):])[:-1].replace('PARTIES:','').strip())               
                if 'interested parties may include' in j and interested_parties==None:
                    interested_parties=j.split(':')[-1]
            for j in extracted_text_1:
                if 'executed by' in j and executor==None:
                    executor=((j[j.index('executed by'):])[:(j[j.index('executed by'):]).index(',')].replace('executed by','').strip())
            substitute_trustee_sale.append([executor,property_address,interested_parties])

        elif "NOTICE OF FORECLOSURE SALE STATE" in i:
            extracted_text=(nltk.tokenize.sent_tokenize(i))
            executor=None
            property_address=None
            interested_parties=None
            for j in extracted_text:
                if 'executed a' in j and executor==None:
                    executor=(((j[j.index('WHEREAS,'):])[:(j[j.index('WHEREAS,'):]).index('executed a')].replace('WHEREAS,','').strip()))
                if 'address' in j.lower() and property_address==None:
                    property_address=(((j[j.index('Address/Description:'):])[:(j[j.index('Address/Description:'):]).index('Current')].replace('Address/Description:','').strip()))
                if 'interested part' in j.lower() and interested_parties==None:
                    interested_parties=(((j[j.index('Interested Party(ies):'):])[:(j[j.index('Interested Party(ies):'):]).index('The sale of the property ')].replace('Interested Party(ies):','').strip()))
            notice_of_foreclosure_sale_state.append([executor,property_address,interested_parties])
    
        else:
            unfiltered_data.append([i])

            
    notice_to_creditors = [list(t) for t in set(tuple(element) for element in notice_to_creditors)]
    notice_to_creditors.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    notice_to_creditors.insert(0, [' - '])
    notice_to_creditors.insert(2, [' - '])
    
    notice_of_successor_trustee_sale = [list(t) for t in set(tuple(element) for element in notice_of_successor_trustee_sale)]
    notice_of_successor_trustee_sale.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    notice_of_successor_trustee_sale.insert(0, [' - '])
    notice_of_successor_trustee_sale.insert(2, [' - '])
    
    substitute_trustee_sale = [list(t) for t in set(tuple(element) for element in substitute_trustee_sale)]
    substitute_trustee_sale.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    substitute_trustee_sale.insert(0, [' - '])
    substitute_trustee_sale.insert(2, [' - '])
    
    notice_of_trustee_sale = [list(t) for t in set(tuple(element) for element in notice_of_trustee_sale)]
    notice_of_trustee_sale.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    notice_of_trustee_sale.insert(0, [' - '])
    notice_of_trustee_sale.insert(2, [' - '])
    
    trustee_sale = [list(t) for t in set(tuple(element) for element in trustee_sale)]
    trustee_sale.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    trustee_sale.insert(0, [' - '])
    trustee_sale.insert(2, [' - '])
    
    substitute_trustee_notice_of_sale = [list(t) for t in set(tuple(element) for element in substitute_trustee_notice_of_sale)]
    substitute_trustee_notice_of_sale.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    substitute_trustee_notice_of_sale.insert(0, [' - '])
    substitute_trustee_notice_of_sale.insert(2, [' - '])
    
    notice_of_substitute_trustee_sale = [list(t) for t in set(tuple(element) for element in notice_of_substitute_trustee_sale)]
    notice_of_substitute_trustee_sale.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    notice_of_substitute_trustee_sale.insert(0, [' - '])
    notice_of_substitute_trustee_sale.insert(2, [' - '])
    
    notice_of_foreclosure_sale_state = [list(t) for t in set(tuple(element) for element in notice_of_foreclosure_sale_state)]
    notice_of_foreclosure_sale_state.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    notice_of_foreclosure_sale_state.insert(0, [' - '])
    notice_of_foreclosure_sale_state.insert(2, [' - '])
    
    unfiltered_data = [list(t) for t in set(tuple(element) for element in unfiltered_data)]
    unfiltered_data.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    unfiltered_data.insert(0, [' - '])
    unfiltered_data.insert(2, [' - '])
        
    try:
        service = build('sheets', 'v4', credentials=creds)

        sheet = service.spreadsheets()

        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range="Davidson Notice to Creditors!A2", valueInputOption="USER_ENTERED", body={"values":notice_to_creditors}).execute()

        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range="Davidson Notice of Successor Trustee Sale!A2", valueInputOption="USER_ENTERED", body={"values":notice_of_successor_trustee_sale}).execute()
        
        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range="Davidson Substitute Trustee Sale!A2", valueInputOption="USER_ENTERED", body={"values":substitute_trustee_sale}).execute()
        
        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range="Davidson Notice of Trustee Sale!A2", valueInputOption="USER_ENTERED", body={"values":notice_of_trustee_sale}).execute()
        
        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range="Davidson Trustee Sale!A2", valueInputOption="USER_ENTERED", body={"values":trustee_sale}).execute()
        
        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range="Davidson Substitute Trustee Notice of Sale!A2", valueInputOption="USER_ENTERED", body={"values":substitute_trustee_notice_of_sale}).execute()
        
        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range="Davidson Notice of Substitute Trustee Sale!A2", valueInputOption="USER_ENTERED", body={"values":notice_of_substitute_trustee_sale}).execute()
        
        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range="Davidson Notice of Foreclosure Sale State!A2", valueInputOption="USER_ENTERED", body={"values":notice_of_foreclosure_sale_state}).execute()
        
        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range="Davidson Unfiltered Data!A2", valueInputOption="USER_ENTERED", body={"values":unfiltered_data}).execute()

    except HttpError as err:
        print(err)
    
        
if county.strip()=='Robertson':
    unmanaged_data=[]
    managed_data=[]
    other_county_data=[]
    
    county='Robertson'
if county.strip()=='Robertson':    
    unmanaged_data=[]
    managed_data=[]
    other_county_data=[]
    count=1
    for i in all_data:
        extracted_text=i.splitlines()
        print(count)
        estate_of_status=False
        estate_of=None
        city=None
        executor=None
        attorney=None
        tn_status=False
        current_data=[None,None,None,None]
        for j in extracted_text:
            if 'Estate of' in j and estate_of==None: 
                estate_of_status=True
                estate_of=j.replace('Estate of','')
                estate_of=estate_of.replace(', Deceased','').strip()
                current_data[1]=(estate_of.strip())
                print('Estate of: ',estate_of)
            if ', TN' in j and city==None :
                if len(j)<90:
                    tn_status=True
                    city=j
                    print(city)
                    current_data[0]=(city.strip())
            if 'administratrix' in j.lower() and executor==None:
                if ',' in j:
                    executor=j.split(',')[0]
                else:
                    executor=extracted_text[extracted_text.index(j)-1]
                current_data[2]=(executor.strip())
            if 'executor' in j.lower()and executor==None:
                if ',' in j:
                    executor=j.split(',')[0]
                else:
                    executor=extracted_text[extracted_text.index(j)-1]
                current_data[2]=(executor.strip())
            if 'executrix' in j.lower()and executor==None:
                if ',' in j:
                    executor=j.split(',')[0]
                else:
                    executor=extracted_text[extracted_text.index(j)-1]
                current_data[2]=(executor.strip())
            if 'personal representative' in j.lower()and executor==None:
                if ',' in j:
                    executor=j.split(',')[0]
                else:
                    executor=extracted_text[extracted_text.index(j)-1]
                current_data[2]=(executor.strip())
            if 'administrator' in j.lower()and executor==None:
                if ',' in j:
                    executor=j.split(',')[0]
                else:
                    executor=extracted_text[extracted_text.index(j)-1]
                current_data[2]=(executor.strip())
            if 'attorney' in j.lower() and attorney==None:
                if ',' in j:
                    attorney=j.split(',')[0]
                else:
                    attorney=extracted_text[extracted_text.index(j)-1]
                current_data[3]=(attorney.strip())
            if None not in current_data:
                break    
        if estate_of_status==False or tn_status==False:
            print('here')
            unmanaged_data.append([i])
        else:
            print(len(current_data))
            managed_data.append(current_data)
        count=count+1
        
    sumner_county_data=[]
    for i in managed_data:
        if 'Sumner' in i[0]:
            sumner_county_data.append(managed_data.pop(managed_data.index(i)))
            
    other_county_data.append(sumner_county_data)
    
    managed_data = [list(t) for t in set(tuple(element) for element in managed_data)]
    managed_data.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    managed_data.insert(0, [' - '])
    managed_data.insert(2, [' - '])
    
    unmanaged_data = [list(t) for t in set(tuple(element) for element in unmanaged_data)]
    unmanaged_data.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    unmanaged_data.insert(0, [' - '])
    unmanaged_data.insert(2, [' - '])
    
    other_county_data = [list(t) for t in set(tuple(element) for element in other_county_data)]
    other_county_data.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    other_county_data.insert(0, [' - '])
    other_county_data.insert(2, [' - '])
    
    try:
        service = build('sheets', 'v4', credentials=creds)

        sheet = service.spreadsheets()

        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range=f"Robertson Managed Data!A2", valueInputOption="USER_ENTERED", body={"values":managed_data}).execute()

        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range=f"Robertson Unmanaged Data!A2", valueInputOption="USER_ENTERED", body={"values":unmanaged_data}).execute()
        
        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range=f"Sumner!A2", valueInputOption="USER_ENTERED", body={"values":other_county_data}).execute()

    except HttpError as err:
        print(err)

if county.strip()=='Wilson':
    managed_data=[]
    for i in all_data:
        extracted_text=i.splitlines()
        estate_of_found=True
        executor='-'
        attorney='-'
        for j in extracted_text:
                if ('estate of' in j.lower()) and (estate_of_found):
                    es=j.replace(':','')
                    es=es.title()
                    estate_of_extracter=es.split('Of')
                    if estate_of_extracter[-1]=='':
                        estate_of=extracted_text[extracted_text.index(j)+1].title()
                        estate_of_found=False
                    if estate_of_extracter[-1]!='':
                        estate_of=es.split('Of')[-1].strip()
                        estate_of_found=False
                if (', deceased' in j.lower()) and (estate_of_found):
                        estate_of=j.split(',')[0].title()
                        estate_of_found=False
                if ('of Property:' in j) and (estate_of_found):
                    estate_of=j.split(':')[-1].strip().title()
                    estate_of_found=False
                if ('OWNER(S):' in j) and (estate_of_found):
                    estate_of=j.split(':')[-1].strip().title()
                    estate_of_found=False 
                if ('PERSONAL' in j.upper()) or ('executor' in j.lower()) or ('administratrix' in j.lower()) or('executrix' in j.lower()) or('administrator' in j.lower()):
                    executor=extracted_text[extracted_text.index(j)-1]
                if 'Attorney' in j or 'ATTORNEY' in j:
                    if ',' in j:
                        attorney=j.split(',')[0]
                    else:
                        attorney=extracted_text[extracted_text.index(j)-1]

        managed_data.append([estate_of.strip(),executor.title(),attorney.title().strip()])
    
    managed_data = [list(t) for t in set(tuple(element) for element in managed_data)]
    managed_data.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    managed_data.insert(0, [' - '])
    managed_data.insert(2, [' - '])
    
    try:
        service = build('sheets', 'v4', credentials=creds)

        sheet = service.spreadsheets()

        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                        range=f"Wilson!A2", valueInputOption="USER_ENTERED", body={"values":managed_data}).execute()
        
    except HttpError as err:
        print(err)
    
    
if county.strip()=='Rutherford':
    
    trustee_sale_data=[]
    substitute_trustee_data=[]
    foreclosure_sale_data=[]
    unfilterable_data=[]
    notice_to_creitors_data=[]
    for i in all_data:
        executed_by='Not available'
        property_address='Not available'
        interested_parties='Not available'
        no_of_parties='Not available'
        property_owners='Not available'
        interested_parties='Not available'
        if 'NOTICE OF TRUSTEE\'S SALE' in i:
            print('NOTICE OF TRUSTEE\'S SALE')
            extracted_text=i.splitlines()
            for j in extracted_text:
                if 'executed by' in j:
                    executed_by=((j[j.index('executed by'):])[:(j[j.index('executed by'):]).index(',')].replace('executed by','').strip())
                    
                if 'ALSO KNOWN AS' in j:
                    property_address=(j.replace('ALSO KNOWN AS:','').strip())
                    
                if ('The sale held pursuant to this Notice' in j):
                    count=1
                    factor=1
                    if 'On or about' in extracted_text[extracted_text.index(j)-1]:
                        count=2
                        factor=2
                    interested_parties=''
                    while True:
                        interested_parties+=(extracted_text[extracted_text.index(j)-count])+'|'
                        count=count+1
                        if 'referenced property:' in extracted_text[extracted_text.index(j)-count]:
                            
                            no_of_parties=count-factor
                            
                            break
            print('Executed by: ',executed_by)
            print('Address: ',property_address)
            print('Interested parties: ',interested_parties)
            print('Number of parties: ',no_of_parties)
            trustee_sale_data.append([executed_by,property_address,interested_parties,no_of_parties])
        elif ('SUBSTITUTE TRUSTEE\'S SALE' in i) or ('SUBSTITUTE TRUSTEES SALE' in i) or ('SUBSTITUTE TRUSTEE\'S NOTICE OF SALE' in i):
            print('SUBSTITUTE TRUSTEE')
            extracted_text=i.splitlines()
            for j in extracted_text:
                if 'executed by' in j:
                    executed_by=((j[j.index('executed by'):])[:(j[j.index('executed by'):]).index(',')].replace('executed by','').strip())
                if ('of Property:' in j) or ('OWNER(S):' in j):
                    property_owners=j[j.index(':'):].replace(':','').strip()                  
                if 'street address' in j.lower():
                    if ':' in j:
                        property_address=j[j.index(':'):].replace(':','').strip()
                    if 'In the event of' in j:
                        print('here')
                        property_address=((j[j.index('is believed to be'):])[:(j[j.index('is believed to be'):]).index('.')].replace('is believed to be','').strip())               
                    if 'but such address' in j:
                        property_address=((j[j.index('is believed to be'):])[:(j[j.index('is believed to be'):]).index(', but such address')].replace('is believed to be','').strip())               
                        print(property_address)
                if 'OTHER INTERESTED PARTIES:' in j.upper():
                    interested_parties=j[j.index(':'):].replace(':','').strip()
                    if interested_parties=='':
                        count=1
                        interested_parties=''
                        while True:
                            interested_parties+=(extracted_text[extracted_text.index(j)+count])+'|'
                            count=count+1
                            if 'The sale of the above-described property' in extracted_text[extracted_text.index(j)+count]:
                                break
                    if 'AM, local time,' in interested_parties:
                        interested_parties=(j[j.index('Other interested parties'):j.index('The hereinafter described')]).replace('The hereinafter described','')
                        interested_parties=interested_parties.replace('Other interested parties','')
                        interested_parties=interested_parties.replace(':','')
                        

            if property_owners=='Not available':
                property_owners=executed_by
            print('Property owner: ',property_owners)
            print('Address: ',property_address)
            print('Interested parties: ',interested_parties)
            substitute_trustee_data.append([property_owners,property_address,interested_parties])
        elif 'NOTICE OF SUBSTITUTE TRUSTEE`S SALE' in i:
            print('NOTICE OF SUBSTITUTE TRUSTEE`S SALE 2nd kind')
            extracted_text=(nltk.tokenize.sent_tokenize(i))
            for j in extracted_text:
                if 'executed by' in j:
                    executed_by=((j[j.index('executed by'):])[:(j[j.index('executed by'):]).index(',')].replace('executed by','').strip())     
                if 'street address' in j.lower():
                    property_address=((j[j.index('Commonly known as'):])[:(j[j.index('Commonly known as'):]).index('The street address')].replace('Commonly known as','').strip())   
                if 'referenced property:' in j.lower():
                    interested_parties=((j[j.index('referenced property:'):])[:(j[j.index('referenced property:'):]).index('.')].replace('referenced property:','').strip())    
            print('Property owner: ',executed_by)
            print('Address: ',property_address)
            print('Interested parties: ',interested_parties)
            substitute_trustee_data.append([executed_by,property_address,interested_parties])
        elif 'NOTICE OF FORECLOSURE SALE' in i.upper():
            print('NOTICE OF FORECLOSURE SALE')
            extracted_text=i.splitlines()
            for j in extracted_text:
                if 'owner(s)' in j.lower():
                    property_owners=j[j.index(':'):].replace(':','').strip() 
                elif 'TRUSTEE' in j:
                    property_owners=extracted_text[extracted_text.index(j)+1].strip()
                if ('address' in j.lower()) and (':' in j):
                    property_address=j[j.index(':'):].replace(':','').strip()
                elif ('address' in j.lower() and 'real estate' in j.lower()):
                    property_address=j[j.index('notice is'):j.index('.')].replace('notice is','').strip()
                if (('interested' in j.lower()) and ('party(ies)' in j.lower()) and (':' in j)) or (('interested' in j.lower()) and ('parties' in j.lower()) and (':' in j)): 
                    interested_parties=j[j.index(':'):].replace(':','').strip()
            print('Property owner: ',property_owners)
            print('Address: ',property_address)
            print('Interested parties: ',interested_parties)
            foreclosure_sale_data.append([property_owners,property_address,interested_parties])
            
        elif 'Notice to Creditors' in i:
            extracted_text=i.splitlines()
            for j in extracted_text:
                if ('Estate of' in j) and len(j)>60:
                    if 'Deceased' in j:
                        property_owners=((j[j.index('Estate of'):])[:(j[j.index('Estate of'):]).index('Deceased')]).replace('Estate of','').strip()
                        property_owners.replace(',','')
                if ('Estate of' in j) and len(j)<=60:
                    property_owners=(j.replace('Estate of','')).strip()
                    property_owners.replace(',','')
                if 'executor' in j.lower():
                    executed_by=(extracted_text[extracted_text.index(j)-1])
                if 'administrator' in j.lower():
                    executed_by=(extracted_text[extracted_text.index(j)-1])
                if 'personal representative' in j.lower():
                    executed_by=(extracted_text[extracted_text.index(j)-1])
                if 'attorney' in j.lower():
                    attorney=(extracted_text[extracted_text.index(j)-1])
            if ',' in property_owners:
                property_owners=property_owners[:-1]
            print('Property owner:',property_owners)
            print('Executors:',executed_by)
            print('Attorney:',attorney)
            notice_to_creitors_data.append([property_owners,executed_by,attorney])
        else:
            print('unfiltered')
            unfilterable_data.append([i])
        print('**************************************')
        
    trustee_sale_data = [list(t) for t in set(tuple(element) for element in trustee_sale_data)]
    trustee_sale_data.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    trustee_sale_data.insert(0, [' - '])
    trustee_sale_data.insert(2, [' - '])
    
    substitute_trustee_data = [list(t) for t in set(tuple(element) for element in substitute_trustee_data)]
    substitute_trustee_data.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    substitute_trustee_data.insert(0, [' - '])
    substitute_trustee_data.insert(2, [' - '])
    
    foreclosure_sale_data = [list(t) for t in set(tuple(element) for element in foreclosure_sale_data)]
    foreclosure_sale_data.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    foreclosure_sale_data.insert(0, [' - '])
    foreclosure_sale_data.insert(2, [' - '])
    
    unfilterable_data = [list(t) for t in set(tuple(element) for element in unfilterable_data)]
    unfilterable_data.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    unfilterable_data.insert(0, [' - '])
    unfilterable_data.insert(2, [' - '])
    
    notice_to_creitors_data = [list(t) for t in set(tuple(element) for element in notice_to_creitors_data)]
    notice_to_creitors_data.insert(0, [f'FROM  {starting_date_entry}  TO   {ending_date_entry}'])
    notice_to_creitors_data.insert(0, [' - '])
    notice_to_creitors_data.insert(2, [' - '])
    
    try:
        service = build('sheets', 'v4', credentials=creds)

        sheet = service.spreadsheets()

        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range=f"Rutherford Trustee Sale!A2", valueInputOption="USER_ENTERED", body={"values":trustee_sale_data}).execute()

        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range=f"Rutherford Substitute Trustee!A2", valueInputOption="USER_ENTERED", body={"values":substitute_trustee_data}).execute()
        
        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range=f"Rutherford Foreclosure Sale!A2", valueInputOption="USER_ENTERED", body={"values":foreclosure_sale_data}).execute()

        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range=f"Rutherford Notice to Creditors!A2", valueInputOption="USER_ENTERED", body={"values":notice_to_creitors_data}).execute()

        request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range=f"Rutherford Unfilterable Data!A2", valueInputOption="USER_ENTERED", body={"values":unfilterable_data}).execute()

    except HttpError as err:
        print(err)
        
driver.close()
