
# coding: utf-8

# In[129]:

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pprint
import pandas as pd
import json
import datetime
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert


# In[130]:
#Default workingDate
workingDate="09/16/2017"
sheetName="your school google attendance sheet name"
#This flag is set to true to run the automation as dry run so it does not save the data
runSelenium = True
delayTimer=2
longDelayTimer=3
saveData=False


# In[131]:

if(runSelenium):
    print ('Automaion will be executed')
else:
    print ('No automaion will be executed')
#if no date is give this will run for last saturday
if(workingDate==""):
    today = datetime.date.today()
    idx = (today.weekday() + 1) % 7
    sat = today - datetime.timedelta(7+idx-6)
    workingDate = '{:%m/%d/%Y}'.format(sat)
"workingDate = "+ workingDate


# In[132]:

scope = ['https://spreadsheets.google.com/feeds']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)
pp = pprint.PrettyPrinter(indent=4)
print ('Plumbing for reading the sheet is done')


# In[133]:

# Extract and process the values
workbook = client.open(sheetName)
sheet1 = workbook.get_worksheet(0)
df=sheet1.get_all_records()
df=pd.DataFrame(df)
k=0
totalData=[]
totalDataAsMap={}
#totalDataAsMap['workingDate']=workingDate
finalDataMap={}
finalDataMap['workingDate']=workingDate
finalDataMap['attDataByClass']=[]
for index, row in df.iterrows():
    if (row['ClassAttendanceSheetName']!=''):
        interSheet=workbook.worksheet(row['ClassAttendanceSheetName'])
        totalData=[]
        if(interSheet):
            print (row['Class']+"---"+row['Name'])
            classData = interSheet.get_all_records()
            totalDataAsMap['Grade']=row['Class']
            totalDataAsMap['Section']=row['Name']
            totalDataAsMap['studentData']=[]
            for student in classData:
                
                if(student[workingDate]=='A'):
                    rowData={}
                    rowData['Student'] = student['Name']
                    rowData['Status'] = student[workingDate]
                    rowData['workingDate']=workingDate
                    rowData['Grade']=row['Class']
                    rowData['Section']=row['Name']
                    #totalDataAsMap['studentData'].append(rowData)
                    totalData.append(rowData)
                totalDataAsMap['studentData']=(totalData)
                k=k+1
            finalDataMap['attDataByClass'].append(totalDataAsMap)
            totalDataAsMap={}


# In[134]:

pp.pprint(finalDataMap['workingDate'])
for cls in finalDataMap['attDataByClass']:
    if(cls['Section'] != ''):
        pp.pprint(cls['Grade']+"="+cls['Section'])
        for student in cls['studentData']:
            pp.pprint("   "+student['Student']+" = "+student['Status'])


# In[135]:

if(runSelenium):
    if(saveData==False):
        print ('Automation suite will not save this')
    driver = webdriver.Chrome('./chromedriver')
    driver.get('https://www.catamilacademy.org/cta/login.aspx');
    time.sleep(delayTimer)  
    email = driver.find_element_by_name('Txt_Mail_Id').send_keys('ctalogin')
    driver.find_element_by_name('Txt_Password').send_keys('ctapassword')
    driver.find_element_by_name('Btn_Login').click()
    time.sleep(delayTimer)
    driver.find_element_by_id('BtnOk').click()
    time.sleep(delayTimer)  
    driver.find_element_by_xpath('//*[@id="logoMenu1_ctl01"]/ul/li[2]/a/span').click()
    time.sleep(delayTimer)
    driver.find_element_by_xpath('//*[@id="logoMenu1_ctl01"]/ul/li[2]/div/ul/li[1]/a/span').click()
    time.sleep(delayTimer)
    driver.switch_to.frame("contentFrame");
    time.sleep(delayTimer)
    for cls in finalDataMap['attDataByClass']:
        driver.find_element_by_xpath("//select[@id='ddlworkingdate']/option[text()='"+finalDataMap['workingDate']+"']").click()
        time.sleep(delayTimer)
        driver.find_element_by_xpath("//select[@name='ddlgrade']/option[text()='"+cls['Grade']+"']").click()
        time.sleep(delayTimer)
        driver.find_element_by_xpath("//select[@name='ddlsection']/option[text()='"+cls['Section']+"']").click()
        time.sleep(delayTimer)
        driver.find_element_by_xpath('//*[@id="btnSrch"]').click()
        time.sleep(delayTimer)
        for student in cls['studentData']:
            print("   "+student['Student']+" = "+student['Status'])
            driver.find_element_by_xpath("//tr/td[contains(.,\""+student['Student']+"\")]/following-sibling::td[1]/*/option[text()='"+student['Status']+"']").click()
            time.sleep(longDelayTimer)
        if(saveData):
            driver.find_element_by_xpath('//*[@id="btnSave"]').click()
            Alert(driver).accept()
        else:
            driver.find_element_by_xpath('//*[@id="btnSaveCancel"]').click()
        time.sleep(delayTimer)
        print(cls['Grade']+" - "+cls['Section']+"'s class data has been processed")
    driver.quit()
else:
    print ('Automation suite is not executed')
