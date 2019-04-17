""" This program is to automate EOD report generation which is sent internally at end of the day """

import xlrd
import shutil,os,sys
import pandas as pd
import getpass,xlsxwriter
import datetime
from selenium import webdriver
from time import sleep
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
temp=2

#We have stored the names of TEMS SD team members in excel sheet(P:\imran-TEMS\NC\LID.xlsx). This excel file has two sheets SD and Offshore.
#SD has list of all the members from TEMS SD which include onshore and offshore
#We are using this sheet to get the name of person from their LID
xls = xlrd.open_workbook(r'P:\imran-TEMS\NC\LID.xlsx')
sheet = xls.sheet_by_index(0)

#Extracting all the LID
lid_list = [sheet.cell_value(row, 0) for row in range(sheet.nrows)]

#Exctracting the names
name_list = [sheet.cell_value(row, 1) for row in range(sheet.nrows)]

#Creating dictionary with LID as key and their relevant name as value
tems = dict(zip(lid_list,name_list))

#Adding chrome to start chrome maximized and to disable the top bar which contains text "This chrome is currently automated"
chrome_options=Options()
chrome_options.add_argument("start-maximized")
chrome_options.add_argument("--disable-infobars")

""" Below function is to navigate to BMC remedy page, run queries and download necessary files """
def bmc():
    global driver
    try:
        driver = webdriver.Chrome(chrome_options=chrome_options,executable_path=r"P:\imran-TEMS\selenium-3.4.3\chromedriver-2.35.exe")
    except:
        driver = webdriver.Chrome(chrome_options=chrome_options,executable_path=r"P:\imran-TEMS\selenium-3.4.3\chromedriver.exe")
    driver.get(r'https://coreweb.prod.itsm.srv.westpac.com.au/arsys/home')
    try:
        cursor = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, "//td[@class='prompttext prompttexterr']"))) # click on applications on the side bar
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
    
    #Below block we are getting date of previous working day in date/Month/year format
    #Note: On Monday, previous working would be Friday
    day = datetime.datetime.today().weekday()
    if day == 0:
        date = (datetime.datetime.now() - datetime.timedelta(days=3)).strftime("%d/%m/%Y")
    else:
        date = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime("%d/%m/%Y")

    #Below is the query query which is use to fetch records of all the incidents which were closed in last 24 hours
    resolved_query= '''\'Owner Group+\'  = \"Testing Environment Management Services\"   AND \'Status*\'   = \"Resolved\" AND (\'Environment\' = \"Test\") AND \'Last Modified Date\'  >\"'''+date+ ''' 10:00:00 PM\"  AND \'Incident Type*\' = \"User Service Restoration\"'''

    #Below is the query which is use to fetch records of all the incidents that are currently opened
    opened_query = '''(\'Owner Group+\' = \"Testing Environment Management Services\"  OR \'Assigned Group*+\'  = \"Testing Environment Management Services\")  AND \'Status*\'  < \"Resolved\" AND \'Status*\' < \"Pending\" AND (\'Environment\' = \"Test\") AND \'Incident Type*\' = \"User Service Restoration\"'''

    #When  clicked on advanced search button, new text area will appear on bottom of the page, So below statement, searching for that text area(search box)
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

    #Once the results are fetched  searching for select all button and scrolling back to the top of the page
    try:
        if(driver.find_element_by_xpath("//*[contains(text(), 'No matches')]")):
            global temp
            temp=1
            # Below statement, scrolling to new search btton and clicking on new search button  on toolbar at top half of the page
            cursor = driver.find_element_by_xpath("(//div[@id='TBnewsearch'])[3]")
            driver.execute_script("arguments[0].scrollIntoView()", cursor)
            ActionChains(driver).move_to_element(cursor).click().perform()

            # When  clicked on new search button, new text area will appear on bottom of the page, So below statement, searching for that text area(search box)
            cursor = driver.find_element_by_xpath("(//textarea[@id='arid1005'])[3]")
            # Since the search box appears at bottom of the page, scrolling down to that element using javascript
            driver.execute_script("arguments[0].scrollIntoView()", cursor)

            # Below statement insert the opened query that was saved above in query variable. We can use send_keys command to do the same but we are using Javascript to insert the opened query since send_keys takes alot of time.
            driver.execute_script(''' arguments[0].value= arguments[1]; ''', cursor, opened_query)

            sleep(3)
            # After entering the query, in the below statement clicking on search button
            cursor = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[@class='f1'][contains(text(),'Search')]")))
            ActionChains(driver).move_to_element(cursor).click().perform()
            sleep(3)

            # Once the results are fetched  searching for select all button and scrolling back to the top of the page
            cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, "//a[@class='SelAll btn btn3d TableBtn'][contains(text(),'Select All')]")))
            driver.execute_script("arguments[0].scrollIntoView()", cursor)
            ActionChains(driver).move_to_element(cursor).click().perform()

            # After clicking select all button  clicking on the report button
            cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, "//a[@class='Rep btn btn3d TableBtn'][contains(text(),'Report')]")))
            ActionChains(driver).move_to_element(cursor).click().perform()

            # as soon as report button is clicked, a new browser window pop's up. In the below statement  switching our driver's control to that newly opened window
            window_before = driver.window_handles[0]
            sleep(5)
            driver.switch_to.window(driver.window_handles[1])

            # Now  selecting TEMS Report from the list of queries that is available
            cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, "//span[@style='padding: 1px 4px;float:left;'][contains(text(),'TEMS Report')]")))
            ActionChains(driver).move_to_element(cursor).click().perform()

            # once we select the query,  clicking on run button
            cursor = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//img[@id='reg_img_93272']")))
            ActionChains(driver).move_to_element(cursor).click().perform()
            sleep(2)

            # After clicked on run, new frame will appear with the result of incidents. In below statements switching to that frame
            # Side Note: Remember we have to always switch to new frame before performing any action on the elements of that frame. So always be on lookout for frame if the code is not working :)
            driver.switch_to_frame(driver.find_element_by_tag_name("iframe"))
            sleep(5)

            # After switching the frame, clicking export icon in the top left corner of the frame. We are using action chains since regular command is not working for some reason
            cursor = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[@id='toolbar']//td[4]/input[1]")))
            ActionChains(driver).move_to_element(cursor).click().perform()
            sleep(3)

            # When clicked on export button a new pop up appears, clicking "Ok" button on that popup
            cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, "//div[@id='exportReportDialogokButton']//input[@type='button']")))
            ActionChains(driver).move_to_element(cursor).click().perform()
            sleep(3)
            driver.close()

            # Once clicked ok, the excel file is generated. So, closing currently opened browser logging out and quiting the driver.
            driver.switch_to.window(window_before)
            cursor = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[@class='f9'][contains(text(),'Logout')]")))
            driver.execute_script("arguments[0].scrollIntoView()", cursor)
            ActionChains(driver).move_to_element(cursor).click().perform()
            sleep(1)
            driver.quit()
    except:
            temp=0
            cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, "//a[@class='SelAll btn btn3d TableBtn'][contains(text(),'Select All')]")))
            driver.execute_script("arguments[0].scrollIntoView()", cursor)
            ActionChains(driver).move_to_element(cursor).click().perform()

            # After clicking select all button  clicking on the report button
            cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, "//a[@class='Rep btn btn3d TableBtn'][contains(text(),'Report')]")))
            ActionChains(driver).move_to_element(cursor).click().perform()

            # as soon as report button is clicked, a new browser window pop's up. In the below statement  switching our driver's control to that newly opened window
            window_before = driver.window_handles[0]
            sleep(5)
            driver.switch_to.window(driver.window_handles[1])

            # Now  selecting TEMS Report Resolved query from the list of queries that is available
            cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, "//span[@style='padding: 1px 4px;float:left;'][contains(text(),'TEMS Report Resolved')]")))
            ActionChains(driver).move_to_element(cursor).click().perform()

            # once we select the query,  clicking on run button
            cursor = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//img[@id='reg_img_93272']")))
            ActionChains(driver).move_to_element(cursor).click().perform()

            # After clicked on run, new frame will appear with the result of incidents. In below statements switching to that frame
            # Side Note: Remember we have to always switch to new frame before performing any action on the elements of that frame. So always be on lookout for frame if the code is not working :)
            driver.switch_to_frame(driver.find_element_by_tag_name("iframe"))
            sleep(5)

            # After switching the frame, clicking export icon in the top left corner of the frame. We are using action chains since regular command is not working for some reason
            cursor = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[@id='toolbar']//td[4]/input[1]")))
            ActionChains(driver).move_to_element(cursor).click().perform()

            # When clicked on export button a new pop up appears, clicking "Ok" button on that popup
            cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, "//div[@id='exportReportDialogokButton']//input[@type='button']")))
            ActionChains(driver).move_to_element(cursor).click().perform()
            sleep(5)
            driver.switch_to.default_content()
            driver.close()

            # Once clicked ok, the excel file is for resolved incidents is generated. So closing currenly poped up browser window and switching the driver control to the parent window to get report for currently opened incidents
            driver.switch_to.window(window_before)
            #Below statement, scrolling to new search btton and clicking on new search button  on toolbar at top half of the page
            cursor = driver.find_element_by_xpath("(//div[@id='TBnewsearch'])[3]")
            driver.execute_script("arguments[0].scrollIntoView()", cursor)
            ActionChains(driver).move_to_element(cursor).click().perform()

             #When  clicked on new search button, new text area will appear on bottom of the page, So below statement, searching for that text area(search box)
            cursor = driver.find_element_by_xpath("(//textarea[@id='arid1005'])[3]")
            # Since the search box appears at bottom of the page, scrolling down to that element using javascript
            driver.execute_script("arguments[0].scrollIntoView()", cursor)


            #Below statement insert the opened query that was saved above in query variable. We can use send_keys command to do the same but we are using Javascript to insert the opened query since send_keys takes alot of time.
            driver.execute_script(''' arguments[0].value= arguments[1]; ''' ,cursor,opened_query)

            sleep(3)
            #After entering the query, in the below statement clicking on search button
            cursor =  WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@class='f1'][contains(text(),'Search')]")))
            ActionChains(driver).move_to_element(cursor).click().perform()
            sleep(3)

            #Once the results are fetched  searching for select all button and scrolling back to the top of the page
            cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[@class='SelAll btn btn3d TableBtn'][contains(text(),'Select All')]")))
            driver.execute_script("arguments[0].scrollIntoView()", cursor)
            ActionChains(driver).move_to_element(cursor).click().perform()

            #After clicking select all button  clicking on the report button
            cursor =  WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[@class='Rep btn btn3d TableBtn'][contains(text(),'Report')]")))
            ActionChains(driver).move_to_element(cursor).click().perform()

            #as soon as report button is clicked, a new browser window pop's up. In the below statement  switching our driver's control to that newly opened window
            window_before = driver.window_handles[0]
            sleep(5)
            driver.switch_to.window(driver.window_handles[1])

            # Now  selecting TEMS Report from the list of queries that is available
            cursor =  WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//span[@style='padding: 1px 4px;float:left;'][contains(text(),'TEMS Report')]")))
            ActionChains(driver).move_to_element(cursor).click().perform()

            #once we select the query,  clicking on run button
            cursor =  WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//img[@id='reg_img_93272']")))
            ActionChains(driver).move_to_element(cursor).click().perform()
            sleep(2)

            #After clicked on run, new frame will appear with the result of incidents. In below statements switching to that frame
            #Side Note: Remember we have to always switch to new frame before performing any action on the elements of that frame. So always be on lookout for frame if the code is not working :)
            driver.switch_to_frame(driver.find_element_by_tag_name("iframe"))
            sleep(5)

            #After switching the frame, clicking export icon in the top left corner of the frame. We are using action chains since regular command is not working for some reason
            cursor=WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@id='toolbar']//td[4]/input[1]")))
            ActionChains(driver).move_to_element(cursor).click().perform()
            sleep(3)

            #When clicked on export button a new pop up appears, clicking "Ok" button on that popup
            cursor=WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@id='exportReportDialogokButton']//input[@type='button']")))
            ActionChains(driver).move_to_element(cursor).click().perform()
            sleep(3)
            driver.close()

            #Once clicked ok, the excel file is generated. So, closing currently opened browser logging out and quiting the driver.
            driver.switch_to.window(window_before)
            cursor=WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//div[@class='f9'][contains(text(),'Logout')]")))
            driver.execute_script("arguments[0].scrollIntoView()", cursor)
            ActionChains(driver).move_to_element(cursor).click().perform()
            sleep(1)
            driver.quit()

    

""" Below function is to create backup of eod excel file which was generated on previous run """
def creating_backup(resolved_path,inpndep_path):
    try:
        os.remove(inpndep_path)        
    except Exception as e:
        pass
    
    try:
        #Taking backup of EOD file
        shutil.copy(r'H:\Documents\EOD\eod.xlsx',r'H:\Documents\EOD\eod - backup.xlsx')
        #Deleting the EOD file which was generated on previous run
        os.remove(r'H:\Documents\EOD\eod.xlsx')        
    except Exception as e:
        pass
    
    try:
        os.remove(resolved_path)        
    except Exception as e:
        pass
    
""" In the below function, we are taking excel sheet which contains incidents which were resolved in last 24 hours and integrating it to new EOD file in resolved sheet """    
def resolved_sheet(path):
    #Opening excel sheet which contains incidents which were resolved in last 24 hours 
    wb= xlrd.open_workbook(path)
    sheet = wb.sheet_by_index(0)

    #Reading all the values of incident number in inc variable. So inc variable will contain list of only incidents number
    inc = [sheet.cell_value(row, 0) for row in range(1,sheet.nrows-2)]

    #In below block we will capture env, the environment is mentioned in cost centre of an incident. So performing string operations and getting env value
    env=[]
    for row in range(1,sheet.nrows-2):
        if len(sheet.cell_value(row, 11).split('/')) <=1:
            env+=[' ']
            continue
        env += [sheet.cell_value(row, 11).split('/')[2]]

    #Extracting description of the incident. 
    desc = [sheet.cell_value(row, 1) for row in range(1,sheet.nrows-2)]
    
    #Extracting submit date of the incident.  
    subdate = [sheet.cell_value(row, 3).split()[0] for row in range(1,sheet.nrows-2)]
    
    #Extracting Assigned to team of the incident.  
    assigned = [sheet.cell_value(row, 5) for row in range(1,sheet.nrows-2)]
    
    #Extracting First Name the requestor of the incident. 
    first_name= [sheet.cell_value(row, 8) for row in range(1,sheet.nrows-2)]
    
    #Extracting second Name the requestor of the incident.  
    second_name = [sheet.cell_value(row,9) for row in range(1,sheet.nrows-2)]
    
    #appending first name and last name of the customer 
    name = [first_name[i]+' '+second_name[i] for i in range(len(first_name))]
    
    #Extracting application of the incident. 
    app =[sheet.cell_value(row, 13) for row in range(1,sheet.nrows-2)]

    #In below block we will extract project details , Project details is mentioned in cost centre of an incident. So performing string operations and getting it's value
    proj=[]
    for row in range(1,sheet.nrows-2):
        if len(sheet.cell_value(row, 11).split('/')) <=1:
            proj+=[' ']
            continue
        proj += [sheet.cell_value(row, 11).split('/')[-1]]

    #Extracting Status of the incident.
    status =[sheet.cell_value(row, 7) for row in range(1,sheet.nrows-2)]

    #In below block we will extract POC and their name. Remember we created tems dictionary in the starting of the program? we will use that dictionary to get the name of the POC by providing the LID which we get from resolved sheet excel
    poc=[]
    for row in range(1,sheet.nrows-2):
        try:
            poc +=[tems[sheet.cell_value(row, 6)]]
        except KeyError:
            poc+=[' ']

    #Extracting Closing Comments of the incident.
    rca =[sheet.cell_value(row, 12) for row in range(1,sheet.nrows-2)]

    #Extracting name of the person who closed the incident.
    closed =[sheet.cell_value(row, 14) for row in range(1,sheet.nrows-2)]

    #Creating dataframe using pandas
    df = pd.DataFrame({'Incident No': inc,'Environment':env,'Description':desc,'Submit Date':subdate,'Currently Assigned to':assigned,
                       'Customer':name,'Applications':app,'Project':proj,'Current Status':status,'TEMS- POC':poc,'RCA':rca,'Closed by':closed})

    #dataframe is never in a particular so we create column order variable, this will be used to order the dataframe
    column_order = ['Incident No','Environment','Description','Submit Date','Currently Assigned to',
                       'Customer','Applications','Project','Current Status','TEMS- POC','RCA','Closed by']

    #applying column order to dataframe
    df=df[column_order]
    return df
    
    
def inprogress_and_dependencies(path):
    #Opening excel sheet which contains incidents which are currently opened
    wb= xlrd.open_workbook(path)
    sheet = wb.sheet_by_index(0)

    #Reading all the values of incident number in inc variable. So inc variable will contain list of only incidents number
    inc = [sheet.cell_value(row, 0) for row in range(1,sheet.nrows-2)]

    #In below block we will capture env, the environment is mentioned in cost centre of an incident. So performing string operations and getting env value
    env=[]
    for row in range(1,sheet.nrows-2):
        if len(sheet.cell_value(row, 11).split('/')) <=1:
            env+=[' ']
            continue
        env += [sheet.cell_value(row, 11).split('/')[2]]

    #Extracting description of the incident.
    desc = [sheet.cell_value(row, 1) for row in range(1,sheet.nrows-2)]

    #Extracting submit date of the incident.
    subdate = [sheet.cell_value(row, 3).split()[0] for row in range(1,sheet.nrows-2)]

    #Extracting Assigned to team of the incident. 
    assigned = [sheet.cell_value(row, 5) for row in range(1,sheet.nrows-2)]

    #Extracting First Name the requestor of the incident.
    first_name= [sheet.cell_value(row, 8) for row in range(1,sheet.nrows-2)]

    #Extracting Second Name the requestor of the incident.
    second_name = [sheet.cell_value(row,9) for row in range(1,sheet.nrows-2)]

    #appending first name and last name of the customer 
    name = [first_name[i]+' '+second_name[i] for i in range(len(first_name))]

    #Extracting Status of the incident.
    status =[sheet.cell_value(row, 7) for row in range(1,sheet.nrows-2)]

    #In below block we will extract POC and their name. Remember we created tems dictionary in the starting of the program? we will use that dictionary to get the name of the POC by providing the LID which we get from opened sheet excel
    poc=[]
    for row in range(1,sheet.nrows-2):
        try:
            poc +=[tems[sheet.cell_value(row, 6)]]
        except KeyError:
            poc+=[' ']

    #RCA for opened incidents has to be filled by the one who is working on EOD. Same is updated in the column as well
    rca=['Need to be filled' for row in range(1,sheet.nrows-2)]

    #Creating dataframe using pandas
    data = pd.DataFrame({'Incident No': inc,'Environment':env,'Description':desc,'Submit Date':subdate,'Currently Assigned to':assigned,
                       'Customer':name,'Current Status':status,'TEMS- POC':poc,'RCA':rca})

    #We are dividng above created dataframe into InProgress(incidents which are assigned to TEMS) and Dependencies(Incidents which are not assigned to TEMS)
    inprog = data[data["Currently Assigned to"]=="Testing Environment Management Services"]
    depend = data[data["Currently Assigned to"]!="Testing Environment Management Services"]

    #For the Dependencies we are updating comments to a generic cooments of "Communicated to (Team Name) team to look into this issue."
    for i in list(depend.index.values):
        depend["RCA"][i] ="Communicated to "+ depend["Currently Assigned to"][i] +" team to look into this issue."

    
    #dataframe is never in a particular so we create column order variable, this will be used to order the dataframe
    column_order = ['Incident No','Environment','Description','Submit Date','Currently Assigned to',
                       'Customer','Current Status','TEMS- POC','RCA']
    #applying column order to dataframe
    inprog=inprog[column_order]
    depend=depend[column_order]
    return inprog,depend

""" In Below function we are creatinf excel sheet with three sheets resolved, inprogress and dependencies. Each of these sheets will be filled with dataframes which were created in above functions """
def creating_excel(df,inprog,depend):
    #Creating writer variable which will write data into the excel file. This creates excel file with name EOD.XLSX
    writer= pd.ExcelWriter(r'H:\Documents\EOD\eod.xlsx',engine='xlsxwriter')

    #Below block we are adding three sheets to EOD.XLSX file
    df.to_excel(writer,sheet_name='resolved',index=False)
    inprog.to_excel(writer,sheet_name='inprogress',index=False)
    depend.to_excel(writer,sheet_name='dependencies',index=False)
    

    #Below block we are applying format to the cells
    workbook  = writer.book
    worksheet1 = writer.sheets['resolved']
    worksheet2 = writer.sheets['inprogress']
    worksheet3 = writer.sheets['dependencies']
    header_format = workbook.add_format({
        'bold': True,
        'align':'center',
        'valign': 'vcenter',
        'text_wrap': True,
        'fg_color': '#ccffff',
        'border': 1})


    for col_num, value in enumerate(df.columns.values):
        worksheet1.write(0, col_num, value, header_format)

    for col_num, value in enumerate(inprog.columns.values):
        worksheet2.write(0, col_num, value, header_format)
        
    for col_num, value in enumerate(depend.columns.values):
        worksheet3.write(0, col_num, value, header_format)
        
    format1 =workbook.add_format({'text_wrap': True,'align':'center','valign': 'vcenter','border': 1})

    worksheet1.set_column('A:L', 20,format1)
    worksheet1.set_column('B:B', 15)
    worksheet1.set_column('C:C', 40)
    worksheet1.set_column('D:D', 15)
    worksheet1.set_column('E:E', 35)
    worksheet1.set_column('K:K', 40)
    worksheet1.set_column('A:L', 20)
    worksheet1.set_row(0,30)

    worksheet2.set_column('A:I', 20,format1)
    worksheet2.set_column('C:C', 45)
    worksheet2.set_column('E:E', 35)
    worksheet2.set_column('I:I', 45)
    worksheet2.set_column('A:I', 20)
    worksheet2.set_row(0,30)

    worksheet3.set_column('A:I', 20,format1)
    worksheet3.set_column('C:C', 45)
    worksheet3.set_column('E:E', 35)
    worksheet3.set_column('I:I', 45)
    worksheet3.set_column('A:I', 20)
    worksheet3.set_row(0,30)

    #Saving and closing the file
    writer.save()
    writer.close()


#Checking if EOD folder exist, if not it creates new one
if os.path.isdir(r"H:\Documents\EOD") == False:
    os.mkdir(r"H:\Documents\EOD")

user_id=getpass.getuser() # getting user Id of the User
resolved_path="C:\\Users\\"+user_id+"\\Downloads\\TEMS20Report20Resolved.xls"
inpndep_path = "C:\\Users\\"+user_id+"\\Downloads\\TEMS5fReport.xls"


def creating_excels(inprog, depend):
    # Creating writer variable which will write data into the excel file. This creates excel file with name EOD.XLSX
    writer = pd.ExcelWriter(r'H:\Documents\EOD\eod.xlsx', engine='xlsxwriter')

    # Below block we are adding three sheets to EOD.XLSX file
    inprog.to_excel(writer, sheet_name='inprogress', index=False)
    depend.to_excel(writer, sheet_name='dependencies', index=False)

    # Below block we are applying format to the cells
    workbook = writer.book
    worksheet2 = writer.sheets['inprogress']
    worksheet3 = writer.sheets['dependencies']
    header_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True,
        'fg_color': '#ccffff',
        'border': 1})

    for col_num, value in enumerate(inprog.columns.values):
        worksheet2.write(0, col_num, value, header_format)

    for col_num, value in enumerate(depend.columns.values):
        worksheet3.write(0, col_num, value, header_format)

    format1 = workbook.add_format({'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})

    worksheet2.set_column('A:I', 20, format1)
    worksheet2.set_column('C:C', 45)
    worksheet2.set_column('E:E', 35)
    worksheet2.set_column('I:I', 45)
    worksheet2.set_column('A:I', 20)
    worksheet2.set_row(0, 30)

    worksheet3.set_column('A:I', 20, format1)
    worksheet3.set_column('C:C', 45)
    worksheet3.set_column('E:E', 35)
    worksheet3.set_column('I:I', 45)
    worksheet3.set_column('A:I', 20)
    worksheet3.set_row(0, 30)

    # Saving and closing the file
    writer.save()
    writer.close()


# Checking if EOD folder exist, if not it creates new one
if os.path.isdir(r"H:\Documents\EOD") == False:
    os.mkdir(r"H:\Documents\EOD")

user_id = getpass.getuser()  # getting user Id of the User
inpndep_path = "C:\\Users\\" + user_id + "\\Downloads\\TEMS5fReport.xls"




creating_backup(resolved_path,inpndep_path) #Calling function to create backup

bmc() # Calling funtion to generate 2 excel files resolved and opened
if(temp==0):
    df = resolved_sheet(resolved_path) #Calling function to get resolved dataframe
    inprog,depend = inprogress_and_dependencies(inpndep_path) # Calling function to get inprog and depend dataframe
    creating_excel(df,inprog,depend) # calling function to create the excel file
else:
    inprog, depend = inprogress_and_dependencies(inpndep_path)  # Calling function to get inprog and depend dataframe
    creating_excels(inprog, depend)  # calling function to create the excel file
print(temp)
print('Excel sheet has been created successfully at H:\Documents\EOD\eod.xlsx. Please make neccessary changes')
