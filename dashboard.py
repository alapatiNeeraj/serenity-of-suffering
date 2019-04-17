import datetime,os,getpass,sys
from selenium import webdriver
from time import sleep
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
from selenium.webdriver.common.action_chains import ActionChains
from xlrd import open_workbook

"""This program is to automate dashboard. We are using python to automate this, we need to install Python, chromedriver. Modules that are required is Selenium, pandas, xlrd,getpass"""

""" Below function is to navigate to BMC remedy page, run queries and download necessary files """
def bmc():
    options = webdriver.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("disable-infobars")
    try:
        driver = webdriver.Chrome(executable_path=r"P:\imran-TEMS\selenium-3.4.3\chromedriver-2.35.exe",chrome_options =options)
    except Exception as e:
        driver = webdriver.Chrome(executable_path=r"P:\imran-TEMS\selenium-3.4.3\chromedriver.exe",chrome_options =options)
    driver.get(r'https://coreweb.prod.itsm.srv.westpac.com.au/arsys/home')
    driver.maximize_window()
    sleep(2)
    try:
        cursor = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, "//td[@class='prompttext prompttexterr']"))) # click on applications on the side bar
        if cursor.text== "User is currently connected from another machine (ARERR 9084)": # if page is not correctly loaded, reload the page
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//div[@class='f9'][contains(text(),'Logout')]"))).click()
            sleep(2)
            driver.quit()
            try:
                driver = webdriver.Chrome(chrome_options=chrome_options,executable_path=r"P:\imran-TEMS\selenium-3.4.3\chromedriver-2.33.exe")
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
    cursor =  WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//span[contains(@class,'navLabel root')][contains(text(),'Incident Management')]"))).click()

    #Below statement   clicking on search incident 
    cursor =  WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//span[contains(@class,"navLabel lvl1")][contains(text(),"Search Incident")]'))).click()
   
    sleep(3)

    #Below statement   clicking on advanced search button  on toolbar at top half of the page
    cursor =  WebDriverWait(driver, 10).until( EC.presence_of_element_located((By.XPATH, "(//div[@id='TBadvancedsearch'])[last()]")))
    cursor.click()
    
    #Below is the query query which is use to fetch records of all the incidents that are currently opened
    query= ''' ('Owner Group+' = "Testing Environment Management Services" OR 'Assigned Group*+' = "Testing Environment Management Services") AND 'Status*' < "Resolved" AND ('Environment' = "Test") '''

    #When  clicked on advanced search button, new text area will appear on bottom of the page, So below statement, searching for that text area(search box)
    cursor = driver.find_element_by_xpath("(//textarea[@id='arid1005'])[3]")
    # Since the search box appears at bottom of the page, scrolling down to that element using javascript
    driver.execute_script("arguments[0].scrollIntoView()", cursor) 

    #Below statement insert the query that was saved above in query variable. We can use send_keys command to do the same but we are using Javascript to insert the query since send_keys takes alot of time.
    driver.execute_script(''' arguments[0].value= arguments[1]; ''' ,cursor,query)

    sleep(1)
    #After entering the query, in the below statement clicking on search button
    cursor =  WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@class='f1'][contains(text(),'Search')]")))
    cursor.click()
    sleep(3)
    #Once the results are fetched  searching for select all button and scrolling back to the top of the page
    cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[@class='SelAll btn btn3d TableBtn']")))
    driver.execute_script("arguments[0].scrollIntoView()", cursor) 
    cursor.click()

    #After clicking select all button  clicking on the report button
    cursor =  WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[@class='Rep btn btn3d TableBtn']")))
    cursor.click()

    #as soon as report button is clicked, a new browser window pop's up. In the below statement  switching our driver's control to that newly opened window
    window_before = driver.window_handles[0]
    sleep(5)
    driver.switch_to.window(driver.window_handles[1])

    # Now  selecting TEMS Report query from the list of queries that is available
    cursor =  WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//span[@style='padding: 1px 4px;float:left;'][contains(text(),'TEMS Report')]")))
    cursor.click()

    #once we select the query,  clicking on run button 
    cursor =  WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//img[@id='reg_img_93272']")))#Run
    cursor.click()
    sleep(7)
    #After clicked on run, new frame will appear with the result of incidents. In below statements switching to that frame
    #Side Note: Remember we have to always switch to new frame before performing any action on the elements of that frame. So always be on lookout for frame if the code is not working :)
    driver.switch_to_frame(driver.find_element_by_tag_name("iframe"))

    #After switching the frame, clicking export icon in the top left corner of the frame. We are using action chains since regular command is not working for some reason
    cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@id='toolbar']//td[4]/input[1]")))#export dialog
    ActionChains(driver).move_to_element(cursor).click().perform()
    sleep(3)
    #When clicked on export button a new pop up appears, clicking "Ok" button on that popup
    cursor=WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@id='exportReportDialogokButton']//input[@type='button']")))# ok Button
    ActionChains(driver).move_to_element(cursor).click().perform()
    sleep(2)
    #Once clicked ok, the excel file is generated. So quiting the driver. 
    driver.quit()


""" After the CSV file is created, below function will be used to upload the same file in admin portal """    
def admin_console(path):
    try:
        browser = webdriver.Chrome(executable_path=r"P:\imran-TEMS\selenium-3.4.3\chromedriver.exe")
    except:
        browser = webdriver.Chrome(executable_path=r"P:\imran-TEMS\selenium-3.4.3\chromedriver-2.33.exe")

    #TEMS portal is very unstable right now so dashboard link keeps changing. To keep this from failing the script, link for the dashboard is saved and read from excel file.
    #Note:  In futue if the dasboard link changes no need to make changes to the code just change the link in the below Excel.

    #Open and reading dashboard link in the below block
    wb=open_workbook(r'P:\imran-TEMS\Dashboard Update\dashboard link.xlsx')
    sheet=wb.sheet_by_index(0)
    link = sheet.cell_value(1,0)
    browser.get(link)
    browser.maximize_window()
    sleep(3)

    #checking if the dashboard is broken by searching Username element in the webpage
    try:
        username = WebDriverWait(browser,5).until(EC.presence_of_element_located((By.ID,"txtuserName")))
    except:
        print("Dashboard Link seems to be broken.Please check with Sagar Dudhedia and update the link if needed to excel sheet located at P:\imran-TEMS\Dashboard Update\dashboard link.xlsx")
        sys.exit(0)

    #If user name is present passing its value
    username.send_keys("admin")

    #Trying to search for password and entering password
    passwd = WebDriverWait(browser,3).until(EC.presence_of_element_located((By.ID,"txtPassword")))
    passwd.send_keys("admin")

    #Below block is to click import button, send the file path.
    WebDriverWait(browser,3).until(EC.presence_of_element_located((By.ID,"btSubmit"))).click()
    sleep(3)
    browser.find_element_by_name("ctl00$ContentPlaceHolder1$fl_Upload").send_keys(path+"TEMS5fReport.csv")
    sleep(4)
    browser.find_element_by_name("ctl00$ContentPlaceHolder1$btnUpload").click()
    sleep(10)
    browser.quit()

""" Below function is to convert XLS file to CSV. Also removing commas(') if they are present since TEMS admin portal gives error if there is comma anywhere """
def xls_to_csv(path):
    data_xls = pd.read_excel(path+"TEMS5fReport.xls", 'Sheet 1', index_col=None)               # reading excel sheet that is generated from BMC using Pandas
    data_xls['Summary*'] = data_xls['Summary*'].str.replace(',', '')                           # removing commas from summary column of the excel
    data_xls['Summary*'] = data_xls['Summary*'].str.replace('\n', ' ')                         # removing New line from the summary column of the excel
    data_xls['Request Cost Centre'] = data_xls['Request Cost Centre'].str.replace(',', '')     # removing commas from Request Cost Centre of the excel
    data_xls['Last Modified Date'] = pd.to_datetime(data_xls['Last Modified Date'])            # Last Modified date contains date in String Format so, Converting it to date format 
    data_xls['Submit Date']=pd.to_datetime(data_xls['Submit Date'])                            # Submit date contains date in String Format so, Converting it to date format 
    data_xls.to_csv(path+"TEMS5fReport.csv",index=False,date_format='%#d/%m/%Y %#I:%M:%S %p')  # Saving the file in CSV Format


user_id=getpass.getuser() # This is used to get user Id of the current user
path = "C:\\Users\\"+user_id+"\\Downloads\\"

''' In the below block  trying to delete any existing excel file whch was previously generated. If we dont remove the below files, new files will be created with name as TEMS5fReport(1).
So as to keep file name consistent we are removing the files before executing the program'''
try:
    os.remove(path+"TEMS5fReport.xls")
except Exception as e:
    pass
try:
    os.remove(path+"TEMS5fReport.csv")
except Exception as e:
    pass

""" Execution of the program starts below. """
bmc() # Calling BMC Funtion
xls_to_csv(path) # Converting xls to csv since TEMS portal accepts only csv file
admin_console(path) # Uploading above generated csv file to admin portal
print("DashBoard Updated Successfully")
