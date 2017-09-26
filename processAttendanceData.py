
# coding: utf-8

# In[217]:

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


# In[218]:

workingDate="09/16/2017"
sheetName="Kalaimagal Tamil School - 2017 Attendance"
dryRun = True
delayTimer=2
attendanceDataComplete= True
longDelayTimer=3
saveData=False


# In[219]:

if(dryRun):
    print ('Automaion will be executed')
else:
    print ('No automaion will be executed')
if(workingDate==""):
    today = datetime.date.today()
    idx = (today.weekday() + 1) % 7
    sat = today - datetime.timedelta(7+idx-6)
    workingDate = '{:%m/%d/%Y}'.format(sat)
"workingDate = "+ workingDate


# In[212]:

scope = ['https://spreadsheets.google.com/feeds']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)
pp = pprint.PrettyPrinter(indent=4)
print ('Plumbing for reading the sheet is done')


# In[221]:

# Extract and process the values
workbook = client.open(sheetName)
sheet1 = workbook.get_worksheet(0)
teachers=sheet1.get_all_records()
df=pd.DataFrame(teachers)
k=0
noOfStudentAbsent=0
absentData=[]
missingInits=[]
attendanceDataComplete= True
totalDataAsMap={}
#totalDataAsMap['workingDate']=workingDate
finalDataMap={}
finalDataMap['workingDate']=workingDate
finalDataMap['attDataByClass']=[]
for index, row in df.iterrows():
    if (row['ClassAttendanceSheetName']!=''):
        interSheet=workbook.worksheet(row['ClassAttendanceSheetName'])
        thisClassAbsentees=[]
        if(interSheet):
            #print ("Starting the process for "+row['Class']+"---"+row['Name'])
            classData = interSheet.get_all_records()
            totalDataAsMap['Grade']=row['Class']
            totalDataAsMap['Section']=row['Name']
            totalDataAsMap['studentData']=[]
            for student in classData:
                #print (student)
                if(student[workingDate]=='A'):
                    noOfStudentAbsent=noOfStudentAbsent+1
                    rowData={}
                    rowData['Student'] = student['Name']
                    rowData['Status'] = student[workingDate]
                    rowData['workingDate']=workingDate
                    rowData['Grade']=row['Class']
                    rowData['Section']=row['Name']
                    thisClassAbsentees.append(rowData)
                    absentData.append(rowData)
                totalDataAsMap['studentData']=(thisClassAbsentees)
                k=k+1
            if (len(totalDataAsMap['studentData'])== 0) :
                #If all students are present we want the tearcher to confirm by giving her initials
                    if(not str(classData[len(classData)-1][workingDate])):
                        missingInits.append({"Teacher":row['ClassAttendanceSheetName'],
                                            "Grade":row['Class']})
                        attendanceDataComplete= False
            finalDataMap['attDataByClass'].append(totalDataAsMap)
            totalDataAsMap={}
absebtdf= pd.DataFrame(absentData)

print("Total number Students processed :"+ str(k-len(teachers)))
print("Total number Students Absent :"+ str(noOfStudentAbsent))
print("Attendance Data Complete : "+ str(attendanceDataComplete))
absentDf


# In[228]:

if len(missingInits)>0:
    print ("Data shows all present but the Teacher initials are missing for following classes : ")
    initmissDF=pd.DataFrame(missingInits)
    display(initmissDF)


# In[140]:

if(dryRun and attendanceDataComplete):
    if(saveData==False):
        print ('Automation suite will not save this')
    driver = webdriver.Chrome('./chromedriver')
    driver.get('https://www.catamilacademy.org/cta/login.aspx');
    time.sleep(delayTimer)  
    email = driver.find_element_by_name('Txt_Mail_Id').send_keys('username')
    driver.find_element_by_name('Txt_Password').send_keys('pwd')
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
    if(dryRun):
        print ('Automation suite is not executed due to dryrun flag')
    else:
        print('Attendence data is incomplete please ask the teacher to put initials if all present')

