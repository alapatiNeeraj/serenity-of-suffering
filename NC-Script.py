from selenium import webdriver
from time import sleep
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
import sys,re
import xlrd,getpass
import shutil,os,sys
import win32com.client as win32
import getpass,xlsxwriter
from datetime import datetime,date,timedelta


#Module for readin LId from excel sheet
#We have stored the names of TEMS SD team members in excel sheet(P:\imran-TEMS\NC\LID.xlsx). This excel file has two sheets SD and Offshore.
#SD has list of all the members from TEMS SD which include onshore and offshore
#We are using this sheet to get the name of person from their LID
#We will be using offshore team LID to send mail only to them
xls = xlrd.open_workbook(r'P:\imran-TEMS\NC\LID.xlsx')
sheet = xls.sheet_by_index(0)
lid_list = [sheet.cell_value(row, 0) for row in range(sheet.nrows)]
name_list = [sheet.cell_value(row, 1) for row in range(sheet.nrows)]
sheet1 = xls.sheet_by_index(1)
offshore_lid = [sheet1.cell_value(row, 0) for row in range(sheet1.nrows)]

result={}


#Adding chrome to start chrome maximized and to disable the top bar which contains text "This chrome is currently automated"
chrome_options=Options()
#chrome_options.add_argument("--headless")
chrome_options.add_argument("start-maximized")
try:
    driver = webdriver.Chrome(chrome_options=chrome_options,executable_path=r"P:\imran-TEMS\selenium-3.4.3\chromedriver-2.35.exe")
except:
    driver = webdriver.Chrome(chrome_options=chrome_options,executable_path=r"P:\imran-TEMS\selenium-3.4.3\chromedriver.exe")

driver.get(r'https://coreweb.prod.itsm.srv.westpac.com.au/arsys/home')
try:
    cursor = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, "//td[@class='prompttext prompttexterr']")))# click on applications on the side bar
    if cursor.text== "User is currently connected from another machine (ARERR 9084)":# if page is not correctly loaded, reload the page
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//div[@class='f9'][contains(text(),'Logout')]"))).click()
        sleep(2)
        driver.quit()
        try:
            driver = webdriver.Chrome(chrome_options=chrome_options,executable_path=r"P:\imran-TEMS\selenium-3.4.3\chromedriver-2.35.exe")
        except:
            driver = webdriver.Chrome(chrome_options=chrome_options,executable_path=r"P:\imran-TEMS\selenium-3.4.3\chromedriver.exe")
        driver.get(r'https://coreweb.prod.itsm.srv.westpac.com.au/arsys/home')
except Exception as e:
    pass

""" Below function is to navigate to BMC remedy page and perform initial steps """
def initialize():
    #click on applications on the side bar
    # Sometime when we open remedy page we get a pop-up, In the below block, trying to check if popup is present and click "Ok" if it is present
    try:
        cursor =  WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "reg_img_304316340"))).click()
        
    except Exception as e:
        WebDriverWait(driver, 3).until(EC.alert_is_present(),
                                       'Timed out waiting for PA creation ' +
                                       'confirmation popup to appear.')

        alert = driver.switch_to.alert
        alert.accept()
        sleep(2)
        cursor =  WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "reg_img_304316340"))).click()
           

    # Below statement  clicking on Incident Management 
    cursor =  WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//span[contains(@class,'navLabel root')][contains(text(),'Incident Management')]")))
    ActionChains(driver).move_to_element(cursor).click().perform()

    #Below statement   clicking on search incident
    cursor =  WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//span[contains(@class,"navLabel lvl1")][contains(text(),"Search Incident")]')))
    ActionChains(driver).move_to_element(cursor).click().perform()
    sleep(3)

    #Below statement   clicking on advanced search button  on toolbar at top half of the page
    cursor =  WebDriverWait(driver, 10).until( EC.presence_of_element_located((By.XPATH, "(//div[@id='TBadvancedsearch'])[last()]")))
    ActionChains(driver).move_to_element(cursor).click().perform()


""" Below function is to check resolved incidents """
def resolved():
    #Below statement, scrolling to new search btton and clicking on new search button  on toolbar at top half of the page
    cursor =  WebDriverWait(driver, 10).until( EC.presence_of_element_located((By.XPATH, "(//div[@id='TBnewsearch'])[last()]")))
    ActionChains(driver).move_to_element(cursor).click().perform()

    #When  clicked on new search button, new text area will appear on bottom of the page, So below statement, searching for that text area(search box)
    cursor = driver.find_element_by_xpath("(//textarea[@id='arid1005'])[3]")
    # Since the search box appears at bottom of the page, scrolling down to that element using javascript
    driver.execute_script("arguments[0].scrollIntoView()", cursor)

    #Below statement insert the resolved query that was saved above in query variable. We can use send_keys command to do the same but we are using Javascript to insert the resolved query since send_keys takes alot of time.
    driver.execute_script(''' arguments[0].value= arguments[1]; ''' ,cursor,resolved_query)

    sleep(1)

    #After entering the query, in the below statement clicking on search button
    cursor =  WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@class='f1'][contains(text(),'Search')]")))
    ActionChains(driver).move_to_element(cursor).click().perform()
    sleep(2)

    #Once we click on search button, we get the result in table. Reading all the elements of the table.
    try:
        cursor =  WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, "//table[@id='T1020']/tbody/tr")))
    except:
        return

    # Iterating through each element to check if it is NC complaint
    for row in range(len(cursor)-1):
        inc_task_action=[]
        resolution_com = []

        #Clicking on the ith incident
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//table[@id='T1020']/tbody/tr[@arrow='"+str(row)+"']"))).click()

        #Getting Currently assigned team name
        team = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//textarea[@id='arid_WIN_3_1000000217']"))).get_attribute("value")

        #Getting currently assigned to person's name
        assignee = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//textarea[@id='arid_WIN_3_1000000218']"))).get_attribute("value")

        inc_no = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//textarea[@id='arid_WIN_3_1000000161']"))).get_attribute("value")      
        inc_date=datetime.strptime(WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//input[@id='arid_WIN_3_3']"))).get_attribute("value").split()[0], '%m/%d/%Y').date()
        sub_id=WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//textarea[@id='arid_WIN_3_2']"))).get_attribute("value")

        #For resolved query any incidents which comes under this query and assigned to TEMS need to be taken care of
        if team == "Testing Environment Management Services":
            try:
                 result[sub_id] += [inc_no+"(Incident was not resolved correctly)"]
            except:
                 result[sub_id] = [inc_no+"(Incident was not resolved correctly)"]
        else:
            print('passed',inc_no)
            continue



def opened():
    #Below statement, scrolling to new search btton and clicking on new search button  on toolbar at top half of the page
    cursor =  WebDriverWait(driver, 10).until( EC.presence_of_element_located((By.XPATH, "(//div[@id='TBnewsearch'])[last()]")))
    ActionChains(driver).move_to_element(cursor).click().perform()

    #When  clicked on new search button, new text area will appear on bottom of the page, So below statement, searching for that text area(search box)
    cursor = driver.find_element_by_xpath("(//textarea[@id='arid1005'])[3]")
    # Since the search box appears at bottom of the page, scrolling down to that element using javascript
    driver.execute_script("arguments[0].scrollIntoView()", cursor)

    #Below statement insert the opened query that was saved above in query variable. We can use send_keys command to do the same but we are using Javascript to insert the opened query since send_keys takes alot of time.
    driver.execute_script(''' arguments[0].value= arguments[1]; ''' ,cursor,opened_query)

    sleep(1)
    #After entering the query, in the below statement clicking on search button
    cursor =  WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@class='f1'][contains(text(),'Search')]")))
    ActionChains(driver).move_to_element(cursor).click().perform()
    sleep(2)

    #Once we click on search button, we get the result in table. Reading all the elements of the table.
    cursor =  WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, "//table[@id='T1020']/tbody/tr")))

    # Iterating through each element to check if it is NC complaint
    for row in range(len(cursor)-1):
        work_detail=[]

        #Clicking on the ith incident
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//table[@id='T1020']/tbody/tr[@arrow='"+str(row)+"']"))).click()

        #Getting Currently assigned team name
        team = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//textarea[@id='arid_WIN_3_1000000217']"))).get_attribute("value")

        #Getting currently assigned to person's name
        assignee = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//textarea[@id='arid_WIN_3_1000000218']"))).get_attribute("value")
        try:
            #Getting work detail on the left side of the incident
            work_detail =WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//table[@id='T301389614']/tbody/tr[2]"))).text.split('\n')

            #Getting type of the work logged. example: General Information, Status Update, Incident Task / Action
            type_inc = work_detail[0]

            #If Attachment is prsent the work detail length is 7, if not it is 6. So we are extracting information based on this
            if len(work_detail) == 6:
                attachment = 0
                submit_date_inc = datetime.strptime(work_detail[3].split()[0], '%m/%d/%Y').date() #Getting submit date of work logged
            else:
                attachment = work_detail[3]
                submit_date_inc = datetime.strptime(work_detail[4].split()[0], '%m/%d/%Y').date()#Getting submit date of work logged
            
            submitter_inc = work_detail[5] #Getting LID of the submitter who logged the work detail
        except:
            pass

        #Getting Incident no
        inc_no = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//textarea[@id='arid_WIN_3_1000000161']"))).get_attribute("value")      

        #Getting incident Submit date
        inc_date=datetime.strptime(WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//input[@id='arid_WIN_3_3']"))).get_attribute("value").split()[0], '%m/%d/%Y').date()

        #Getting incident submitter ID
        sub_id=WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//textarea[@id='arid_WIN_3_2']"))).get_attribute("value")
        if team == "Test Data Management":
            print("TDM", inc_no)
            continue
        else:
            if len(work_detail) == 0:                         
                if inc_date == date.today() or (date.today() - inc_date).days <2:
                    print("Passed ", inc_no)
                    continue
                else:
                   print("Failed",inc_no)
                   try:
                        result[sub_id] += [inc_no+"(No update since 48 Hrs)"]
                   except:
                        result[sub_id] = [inc_no+"(No update since 48 Hrs)"]
            else:
                if (date.today() - submit_date_inc).days <2 :
                    if submitter_inc in lid_list:
                        if type_inc in ['Incident Task / Action','Status Update']  or (type_inc in ['General Information'] and int(attachment) > 0) or (type_inc in ['General Information'] and inc_date == submit_date_inc):
                            if (team == "Testing Environment Management Services" and len(assignee) > 0) or (team != "Testing Environment Management Services"):
                                print('passed',inc_no)
                                continue
                            else:
                                #print("Team is incorrect and no assigned",inc_no)
                                try:
                                    result[sub_id] += [inc_no+"(Incident Assigned to TEMS but not assigned to individual)"]
                                except:
                                    result[sub_id] = [inc_no+"(Incident Assigned to TEMS but not assigned to individual)"]
                                
                        else:
                            #print('inc type is incorrect',inc_no)
                            try:
                                result[sub_id] += [inc_no+"(Incident Type is incorrect in Work detail.)"]
                            except:
                                result[sub_id] = [inc_no+"(Incident Type is incorrect in Work detail.)"]
                    else:
                        print('passed',inc_no)
                        continue
                else:
                    #print("Not updated since two days",inc_no)
                    try:
                        result[sub_id] += [inc_no+"(No update since 48 Hrs)"]
                    except:
                        result[sub_id] = [inc_no+"(No update since 48 Hrs)"]


def pending():
    #Below statement, scrolling to new search btton and clicking on new search button  on toolbar at top half of the page
    cursor =  WebDriverWait(driver, 10).until( EC.presence_of_element_located((By.XPATH, "(//div[@id='TBnewsearch'])[last()]")))
    ActionChains(driver).move_to_element(cursor).click().perform()

    #When  clicked on new search button, new text area will appear on bottom of the page, So below statement, searching for that text area(search box)
    cursor = driver.find_element_by_xpath("(//textarea[@id='arid1005'])[3]")
    # Since the search box appears at bottom of the page, scrolling down to that element using javascript
    driver.execute_script("arguments[0].scrollIntoView()", cursor)

    #Below statement insert the pending query that was saved above in query variable. We can use send_keys command to do the same but we are using Javascript to insert the pending query since send_keys takes alot of time.
    driver.execute_script(''' arguments[0].value= arguments[1]; ''' ,cursor,pending_query)

    sleep(1)
    #After entering the query, in the below statement clicking on search button
    cursor =  WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@class='f1'][contains(text(),'Search')]")))
    ActionChains(driver).move_to_element(cursor).click().perform()
    sleep(2)
    
    #Once we click on search button, we get the result in table. Reading all the elements of the table.
    cursor =  WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, "//table[@id='T1020']/tbody/tr")))

    # Iterating through each element to check if it is NC complaint
    for row in range(len(cursor)-1):
        work_detail=[]

        #Clicking on the ith incident
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//table[@id='T1020']/tbody/tr[@arrow='"+str(row)+"']"))).click()

        #Getting Currently assigned team name
        team = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//textarea[@id='arid_WIN_3_1000000217']"))).get_attribute("value")
        
        #Getting currently assigned to person's name
        assignee = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//textarea[@id='arid_WIN_3_1000000218']"))).get_attribute("value")

        #Getting target date that is set for the incident 
        target_date = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//input[@id='arid_WIN_3_1000005261']")))

        #Checking if target date is empty
        if len(target_date.get_attribute("value")) > 0:
            target_date = datetime.strptime(target_date.get_attribute("value").split()[0], '%m/%d/%Y').date()
        else:
            target_date=""
            
        try:
            #Getting work detail on the left side of the incident
            work_detail =WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//table[@id='T301389614']/tbody/tr[2]"))).text.split('\n')
            #Getting type of the work logged. example: General Information, Status Update, Incident Task / Action
            type_inc = work_detail[0]
            
            attachment = work_detail[3]
            
            submit_date_inc = datetime.strptime(work_detail[4].split()[0], '%m/%d/%Y').date()#Getting submit date of work logged
            
            submitter_inc = work_detail[5]#Getting LID of the submitter who logged the work detail
        except:
            pass

        #Getting Incident no
        inc_no = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//textarea[@id='arid_WIN_3_1000000161']"))).get_attribute("value")
        
        #Getting incident Submit date
        inc_date=datetime.strptime(WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//input[@id='arid_WIN_3_3']"))).get_attribute("value").split()[0], '%m/%d/%Y').date()

        #Getting incident Submitter ID
        sub_id=WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//textarea[@id='arid_WIN_3_2']"))).get_attribute("value")
        if team == "Test Data Management":
            print("TDM", inc_no)
            continue
        else:
            if len(str(target_date)) == 0:
                if len(work_detail) == 0:
                    try:
                        result[sub_id] += [inc_no+"(Incident is in pending state and no status update and No target date assigned)"]
                    except:
                        result[sub_id] = [inc_no+"(Incident is in pending state and no status update provided and No target date assigned)"]
                else:
                    if type_inc in ['Status Update'] and (date.today() - submit_date_inc).days <=1:
                        print('passed', inc_no)
                        continue
                    else:
                        #print("Not updated since one day",inc_no)
                        try:
                            result[sub_id] += [inc_no+"(Incident is in pending state and not updated since 24 hrs and No target date assigned)"]
                        except:
                            result[sub_id] = [inc_no+"(Incident is in pending state and not updated since 24 hrs and No target date assigned)"]
                            
            elif len(str(target_date)) > 0 and (target_date - date.today()).days >0:
                if len(work_detail) == 0:
                    try:
                        result[sub_id] += [inc_no+"(Target date has been set but no status update provided in work info)"]
                    except:
                        result[sub_id] = [inc_no+"(Target date has been set but no status update provided in work info)"]
                else:
                    if type_inc in ['Status Update']:
                        print('passed', inc_no)
                        continue
                    else:
                        #print("Status update is not provided in work info",inc_no)
                        try:
                            result[sub_id] += [inc_no+"(Status update is not provided in work info)"]
                        except:
                            result[sub_id] = [inc_no+"(Status update is not provided in work info)"]
                        
            else:
                if len(work_detail) == 0:
                    try:
                        result[sub_id] += [inc_no+"(Target date has already been breached and no status update provided)"]
                    except:
                        result[sub_id] = [inc_no+"(Target date has already been breached and no status update provided)"]
                else:
                    if type_inc in ['Status Update'] and (date.today() - submit_date_inc).days <=1:
                        print('passed', inc_no)
                        continue
                    else:
                        #print("Not updated since one day",inc_no)
                        try:
                            result[sub_id] += [inc_no+"(Target date has already been breached and not updated since 24 hrs)"]
                        except:
                            result[sub_id] = [inc_no+"(Target date has already been breached and not updated since 24 hrs)"]


                
""" Below Function is to send mail to team members"""
def sendmail():
    remaining = []
    outlook = win32.Dispatch('outlook.application')
    user_id=getpass.getuser()
    for key,val in result.items():
        if key in offshore_lid:
            mail = outlook.CreateItem(0)
            mail.To = key
            mail.BCC = user_id
            mail.Subject = "Incidents which needs to be checked for NC Report"
            mail.Body = '''Hi '''+name_list[lid_list.index(key)]+''',
            Please Look into below incidents, they come under NC Report. Please update them ASAP.

            Incidents  = '''+str(val)+'''.'''
            mail.send
        else:
            remaining +=val
    mail = outlook.CreateItem(0)
    mail.To = 'temsservicedesk@westpac.com.au'
    mail.BCC = user_id
    mail.Subject = "Incidents which needs to be checked for NC Report"
    mail.Body = '''Hi Team,
    Below incident were either raised by onshore team or someone other than TEMS. Please look into these and update them ASAP.

    Incidents  = '''+str(remaining)+'''.'''
    mail.send


            
day = datetime.today().weekday()
if day==0:
    dates=(datetime.now()-timedelta(days=3)).strftime("%m/%d/%Y")
else:
    dates=(datetime.now()-timedelta(days=1)).strftime("%m/%d/%Y")
resolved_query= '''('Owner Group+'  = "Testing Environment Management Services"  OR 'Assigned Group*+'  = "Testing Environment Management Services")  AND ('Assigned Group*+'  != "Test Data Management") AND 'Status*'  = "Resolved" AND 'Resolution'  = $NULL$ '''
opened_query = '''(\'Owner Group+\' = \"Testing Environment Management Services\"  OR \'Assigned Group*+\'  = \"Testing Environment Management Services\")  AND ('Assigned Group*+'  != "Test Data Management")  AND \'Status*\'  < \"Resolved\" AND 'Status*' < "Pending"'''
pending_query = '''(\'Owner Group+\' = \"Testing Environment Management Services\"  OR \'Assigned Group*+\'  = \"Testing Environment Management Services\")  AND 'Status*' = "Pending"'''

initialize()
resolved()
opened()
pending()
driver.quit()
sendmail()


    
