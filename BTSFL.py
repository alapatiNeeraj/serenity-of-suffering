from selenium import webdriver
from time import sleep
import time,os
from selenium.webdriver.support.ui import WebDriverWait,Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.proxy import *
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.chrome.options import Options
import win32com.client as win32
from datetime import date,timedelta
import getpass,traceback
import logging
import binascii,ctypes
from xlrd import open_workbook
import update_excel as ud
from datetime import datetime
from win32com.client import Dispatch
from win32com.client.gencache import EnsureDispatch
from win32com.client import constants
import re
from bs4 import BeautifulSoup as bs


wb = open_workbook(r"P:\imran-TEMS\BTFSL\BT Report Test Data.xlsx")
sheet =wb.sheet_by_name("Test Data")


env_dom= {'SIT1':'id("DifferentEnvUrlsDropDown")/option[12]','SIT2':'id("DifferentEnvUrlsDropDown")/option[16]','SIT3':'id("DifferentEnvUrlsDropDown")/option[20]','UAT1':'id("DifferentEnvUrlsDropDown")/option[31]','SVP':'id("DifferentEnvUrlsDropDown")/option[27]',
              'SIT1_SAML_FF':'id("DifferentEnvUrlsDropDown")/option[14]','SIT2_SAML_FF':'id("DifferentEnvUrlsDropDown")/option[18]','SIT3_SAML_FF':'id("DifferentEnvUrlsDropDown")/option[22]','UAT1_SAML_FF':'id("DifferentEnvUrlsDropDown")/option[33]',
              'SIT1_NTLM_RB':'id("DifferentEnvUrlsDropDown")/option[13]','SIT2_NTLM_RB':'id("DifferentEnvUrlsDropDown")/option[17]','SIT3_NTLM_RB':'id("DifferentEnvUrlsDropDown")/option[21]','UAT1_NTLM_RB':'id("DifferentEnvUrlsDropDown")/option[32]'}
count=0
chrome_options = Options()
#chrome_options.add_argument("--headless")
exc_path = r"P:\imran-TEMS\BTFSL\BTSFL\chromedriver.exe"
exc_path1 = r"P:\imran-TEMS\BTFSL\BTSFL\chromedriver-2.33.exe"
exc_path_ie=r"P:\imran-TEMS\BTFSL\BTSFL\IEDriverServer.exe"
def btstg(gcis,btno,env):
    global count
    try:
        try:
            #driver = webdriver.Chrome(chrome_options=chrome_options,executable_path=exc_path1)
            driver = webdriver.Ie(exc_path_ie)
        except:
            driver = webdriver.Chrome(chrome_options=chrome_options,executable_path=exc_path)
        logging.info('\n\n')
        logging.info('internet explorer started')

        driver.get("http://dwwas0004.btfin.com:9081/ingress/access")
        driver.maximize_window()
        
        logging.info('Ingress page opened')
        
        cursor = driver.find_element_by_id('site-BTI')
        cursor.click()

        cursor = driver.find_element_by_id('brand-STG')
        cursor.click()

        cursor = driver.find_element_by_id('CustAcctNo')
        cursor.send_keys(gcis)

        cursor=driver.find_element_by_id('acctId')
        cursor.send_keys(btno)

        cursor=driver.find_element_by_id('posted')
        cursor.click()
        
        cursor=driver.find_element_by_xpath(env_dom[env])
        cursor.click()

        cursor = driver.find_element_by_id('generateSAML')
        sleep(1)
        cursor.click()
        logging.info('Ingress page completed')
        sleep(2)
        if driver.title == "This page can’t be displayed" or driver.title == "HTTP 404 Not Found":
            raise Exception()
        try:
            WebDriverWait(driver,3).until(EC.presence_of_element_located((By.XPATH,'//a[contains(text(),"BT Super for Life - Savings")]'))).click()
        except Exception:
            pass

        cursor = WebDriverWait(driver,30).until(EC.presence_of_element_located((By.XPATH,'//h3[contains(text(),"BT Super for Life")] | id("open-my-account") | //article[@class="unauthenticated-text"] | //h1[contains(text(),"Overview")]')))
        if cursor.text == "Sorry, The page you have requested is only available after you have signed in to Internet Banking.":
            raise Exception()
        count=0
        return 'pass'
    
    except Exception as e:
        '''if count <2:
            count =count+1
            driver.close()
            driver.quit()
            res=btstg(gcis,btno,env)
        count=0'''
        '''try:
            return res
        except Exception:
            return 'fail' '''
        return 'fail'
       
    finally:
        try:
            driver.close()
            driver.quit()
        except:
            pass
        

def btwbc(can,env):
    global count
    try:
        try:
            #driver = webdriver.Chrome(chrome_options=chrome_options,executable_path=exc_path)
            driver = webdriver.Ie(exc_path_ie)
        except:
            driver = webdriver.Chrome(chrome_options=chrome_options,executable_path=exc_path1)
        logging.info('\n\n')
        logging.info('internet explorer started')

        driver.get("http://dwwas0004.btfin.com:9081/ingress/access")
        driver.maximize_window()
        
        cursor = driver.find_element_by_id('site-BTI')
        cursor.click()

        cursor = driver.find_element_by_id('CustAcctNo')
        cursor.send_keys(can)


        cursor=driver.find_element_by_id('posted')
        cursor.click()
        
        cursor=driver.find_element_by_xpath(env_dom[env])
        cursor.click()

        cursor = driver.find_element_by_id('generateSAML')
        cursor.click()
        sleep(2)
        if driver.title == "This page can’t be displayed" or driver.title == "HTTP 404 Not Found":
            raise Exception()
        try:
            WebDriverWait(driver,5).until(EC.presence_of_element_located((By.XPATH,'//a[contains(text(),"BT Super for Life")]'))).click()
        except Exception:
            pass
        if 'SIT2' in env:
            return 'pass'
        else:
            cursor = WebDriverWait(driver,30).until(EC.presence_of_element_located((By.XPATH,'//span[contains(text(),"BT Super for Life")] | //h3[contains(text(),"BT Super for Life")] | //input[@id="submitButton"] | //article[@class="unauthenticated-text"]  | //h1[contains(text(),"Overview")]')))
            #cursor = WebDriverWait(driver,30).until(EC.presence_of_element_located((By.XPATH,'//span[contains(text(),"BT Super for Life")]')))
            if cursor.text == "Sorry, The page you have requested is only available after you have signed in to Internet Banking.":
                    raise Exception()          
            count=0
            return 'pass'
    except Exception as e:
        return 'fail'
    finally:
        try:
            driver.close()
            driver.quit()
        except:
            pass


def compass_desktop(can,scn,intp,url):
    global count
    try:
        try:
            #driver = webdriver.Chrome(chrome_options=chrome_options,executable_path=exc_path)
            caps=DesiredCapabilities.INTERNETEXPLORER
            caps['initialBrowserUrl'] =url
            driver = webdriver.Ie(executable_path=exc_path_ie,capabilities=caps)
        except:
            driver = webdriver.Chrome(chrome_options=chrome_options,executable_path=exc_path1)
            driver.get(url)
        logging.info('\n\n')
        logging.info('internet explorer started')

        
        driver.maximize_window()

        cursor = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.ID,'access-number')))
        cursor.click()
        cursor.send_keys(can)
        
        cursor = WebDriverWait(driver,3).until(EC.presence_of_element_located((By.ID,'securityNumber')))
        cursor.click()
        cursor.send_keys(scn)

        cursor = WebDriverWait(driver,3).until(EC.presence_of_element_located((By.ID,'internet-password')))
        cursor.click()
        cursor.send_keys(intp)

        cursor = WebDriverWait(driver,3).until(EC.presence_of_element_located((By.ID,'logonButton')))
        cursor.send_keys(Keys.RETURN)
        try:
            cursor = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'(//a[contains(text(),"BT Super")])[1]')))
            cursor.click()            
        except Exception as e:
            return 'Fail'
        sleep(2)
        if driver.title == "This page can’t be displayed" or driver.title == "HTTP 404 Not Found":
            raise Exception()
        cursor = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//div[contains(text(),"Open account")] |//div[contains(text(),"Back to My Accounts")] | //h3[contains(text(),"BT Super for Life")] | //article[@class="unauthenticated-text"]| //h1[contains(text(),"Overview")]')))
        #driver.execute_script("arguments[0].scrollIntoView()", cursor)
        if cursor.text == "Sorry, The page you have requested is only available after you have signed in to Internet Banking.":
            raise Exception()
        count=0
        return 'pass'
        
    except Exception as e:
        return 'fail'
    finally:
        try:
            driver.close()
            driver.quit()
        except:
            pass



def compass_mobile(can,scn,intp,url):
    global count
    try:
        caps=DesiredCapabilities.INTERNETEXPLORER
        caps['initialBrowserUrl'] =url
        driver = webdriver.Ie(executable_path=exc_path_ie,capabilities=caps)
        #try:
            #driver = webdriver.Chrome(chrome_options=chrome_options,executable_path=exc_path)
        #except:
            #driver = webdriver.Chrome(chrome_options=chrome_options,executable_path=exc_path1)
        logging.info('\n\n')
        logging.info('internet explorer started')

        driver.maximize_window()

        cursor = WebDriverWait(driver,5).until(EC.presence_of_element_located((By.ID,'card-access-no')))
        cursor.click()
        cursor.send_keys(can)
        
        cursor = WebDriverWait(driver,3).until(EC.presence_of_element_located((By.ID,'security-no')))
        cursor.click()
        cursor.send_keys(scn)
        sleep(1)

        cursor = WebDriverWait(driver,3).until(EC.presence_of_element_located((By.ID,'internet-pwd')))
        cursor.click()
        cursor.send_keys(intp)

        sleep(1)
        cursor.send_keys(Keys.RETURN)
        try:
            try:
                cursor = WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH,'//label[@for="accept-terms"]//span')))
                ActionChains(driver).move_to_element(cursor).click().perform()
                cursor = WebDriverWait(driver,3).until(EC.element_to_be_clickable((By.XPATH,'//button[@data-bind="click: vm.logon"]')))
                cursor.click()
                
            except Exception as e:
                #traceback.print_exc()
                pass

            cursor = WebDriverWait(driver,3).until(EC.presence_of_element_located((By.XPATH,'//div[@id="slide-1"]'))).click()
            ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//div[@id="slide-2"]')))).click().perform()
            ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//div[@id="slide-3"]')))).click().perform()
            cursor = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//button[@role="button"]'))).click()

        except Exception as e:
            pass          

        try:
            for i in range(10):
                xpath = 'id("account-tile-'+str(i)+'")//div[@aria-labelledby = "account-description-'+str(i)+'"]//small[@class="account-name"]'
                cursor = WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH,xpath)))
                if 'BT Super' in cursor.text:
                    cursor.click()
                    try:
                        driver.switch_to.window(driver.window_handles[1])
                        driver.maximize_window()
                    except Exception as e:
                        try:
                            cursor = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,xpath)))
                            cursor.send_keys(Keys.CONTROL + Keys.RETURN)
                        except Exception as e:
                            continue
                    break
        except Exception as e:
            count=0
            return 'fail'      
        sleep(2)
        if driver.title == "This page can’t be displayed" or driver.title == "HTTP 404 Not Found":
            raise Exception()

        cursor = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//button[@id="open-my-account"] |//div[contains(text(),"Exit")] | //button[@id="empContrib2"] | //div[@class="col-sm-12 col-md-12 col-lg-12 visible-sm visible-md visible-lg"]| //h1[contains(text(),"Overview")]')))
    
        if cursor.text == "BT Super for Life is currently unavailable. Please come back later or call BT Customer Relations on 1300 653 553 if your enquiry is urgent.":
            raise Exception
        
        driver.execute_script("arguments[0].scrollIntoView()", cursor)
        count=0
        return 'pass'
    except Exception as e:
        return 'fail'
    finally:
        try:
            driver.close()
            driver.quit()
        except:
            pass

def firefly():
    try:
        print("FireFly ")
        ind =['username','password','customer','link']
        firefly_input =dict(zip(ind,[sheet.cell_value(row,3).strip() for row in (25,26,27,28)]))
        driver = webdriver.Ie(executable_path=exc_path_ie)
        driver.get(firefly_input['link'])
        driver.maximize_window()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//input[@type='text']"))).send_keys(firefly_input['username'])
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//input[@type='password']"))).send_keys(firefly_input['password'])
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//input[@alt='Login']"))).click()
        sleep(2)
        cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//frame[@name='DetailsFrm1' and @id='DetailsFrm1']")))
        driver.switch_to.frame(cursor)
        cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//frame[@name='fraLeftFrame']")))
        driver.switch_to.frame(cursor)
        cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//frame[@name='fraDeeptree']")))
        driver.switch_to.frame(cursor)
        num_windows = driver.window_handles
        WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//span[@title='Firefly E2E']//a[contains(text(),'Firefly E2E')]"))).click()
        WebDriverWait(driver,15).until(lambda driver:  len(num_windows) != len(driver.window_handles))
        driver.switch_to.window(driver.window_handles[1])
        driver.maximize_window()
        try:
            ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@id='SGB_SRVC_HS_WRK_SGB_GCIS_ID']")))).click().send_keys(firefly_input['customer']).send_keys(Keys.RETURN).perform()
        except:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//input[@type='text']"))).send_keys(firefly_input['username'])
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//input[@type='password']"))).send_keys(firefly_input['password'])
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,"//input[@alt='Login']"))).click()
            ActionChains(driver).move_to_element(WebDriverWait(driver,15).until(EC.presence_of_element_located((By.XPATH,"//input[@id='SGB_SRVC_HS_WRK_SGB_GCIS_ID']")))).click().send_keys(firefly_input['customer']).send_keys(Keys.RETURN).perform()
            sleep(2)
        cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//frame[@name='UniversalHeader']")))
        driver.switch_to.frame(cursor)
        cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//frame[@name='TargetContent']")))
        driver.switch_to.frame(cursor)
        elements = WebDriverWait(driver,15).until(EC.presence_of_all_elements_located((By.XPATH,"//table[@id='SGB_CUST_ACT_VW$scroll$0']/tbody/tr")))
        num_windows = driver.window_handles
        for i,elem in enumerate(elements[2:]):
            if 'BTSFL' in elem.text:            
                WebDriverWait(driver,15).until(EC.presence_of_element_located((By.XPATH,"//table[@id='SGB_CUST_ACT_VW$scroll$0']//a[@id='FIN_ACCOUNT_ID$"+str(i)+"']"))).click()            
                break
        sleep(10)
        driver.switch_to.window(driver.window_handles[2])
        driver.maximize_window()
        sleep(2)
        if driver.title == "This page can’t be displayed" or driver.title == "HTTP 404 Not Found":
            raise Exception()
        cursor = WebDriverWait(driver,30).until(EC.presence_of_element_located((By.XPATH,'//h3[contains(text(),"BT Super")] | //article[@class="unauthenticated-text"]')))
        if cursor.text == "Sorry, The page you have requested is only available after you have signed in to Internet Banking.":
                raise Exception()
        ud.update([32],[8,9,10,11,12,13],'pass',"BT")
        return 'pass'
    except Exception as e:
        #traceback.print_exc
        ud.update([32],[8,9,10,11,12,13],'fail',"BT")
        return 'fail'
    finally:
        try:
            driver.close()
            driver.quit()
        except:
            pass
    
def wlive(can,pswd,url):
    global count
    try:
        caps=DesiredCapabilities.INTERNETEXPLORER
        caps['initialBrowserUrl'] =url
        driver = webdriver.Ie(executable_path=exc_path_ie,capabilities=caps)
        #try:
            #driver = webdriver.Chrome(chrome_options=chrome_options,executable_path=exc_path)
        #except:
            #driver = webdriver.Chrome(chrome_options=chrome_options,executable_path=exc_path1)
        logging.info('\n\n')
        logging.info('internet explorer started')

        driver.get(url)
        driver.maximize_window()

        cursor = WebDriverWait(driver,5).until(EC.presence_of_element_located((By.XPATH,"//input[@id='fakeusername']"))).send_keys(can)
        #ActionChains(driver).move_to_element(WebDriverWait(driver,3).until(EC.presence_of_element_located((By.XPATH,"//input[@id='password']")))).click().send_keys(pswd).perform()
        try:
            WebDriverWait(driver,3).until(EC.presence_of_element_located((By.XPATH,"//button[contains(text(),'I')]"))).click()
            WebDriverWait(driver,3).until(EC.presence_of_element_located((By.XPATH,"//button[contains(text(),'N')]"))).click()
            WebDriverWait(driver,3).until(EC.presence_of_element_located((By.XPATH,"//button[contains(text(),'T')]"))).click()
            WebDriverWait(driver,3).until(EC.presence_of_element_located((By.XPATH,"//button[contains(text(),'B')]"))).click()
            WebDriverWait(driver,3).until(EC.presence_of_element_located((By.XPATH,"//button[contains(text(),'K')]"))).click()
            WebDriverWait(driver,3).until(EC.presence_of_element_located((By.XPATH,"//button[contains(text(),'1')]"))).click()
        except Exception as e:
            ActionChains(driver).move_to_element(WebDriverWait(driver,3).until(EC.presence_of_element_located((By.XPATH,"//input[@id='password']")))).click().send_keys(pswd).perform()
        
        WebDriverWait(driver,3).until(EC.presence_of_element_located((By.XPATH,"//button[@id='signin']"))).click()
        sleep(3)
        ActionChains(driver).move_to_element(WebDriverWait(driver,15).until(EC.presence_of_element_located((By.XPATH,"//h2[contains(text(),'BT Super')]")))).click().perform()
        if  url == "https://uat.banking.westpac.com.au/":
            return 'pass'
        sleep(5)
        if driver.title == "This page can’t be displayed" or driver.title == "HTTP 404 Not Found":
            raise Exception()
        sleep(5)
        cursor = WebDriverWait(driver,30).until(EC.presence_of_element_located((By.XPATH,'//h3[contains(text(),"BT Super")] | //article[@class="unauthenticated-text"]| //h1[contains(text(),"Overview")] | //div[contains(text(),"Your account is currently inactive")]')))
        if cursor.text == "Sorry, The page you have requested is only available after you have signed in to Internet Banking.":
                raise Exception()
        count=0
        return 'pass'        
    except Exception as e:
        return 'fail'
    finally:
        try:
            driver.close()
            driver.quit()
        except:
            pass



def rb(cust,userid,pswd,url):
    global count
    try:
        caps=DesiredCapabilities.INTERNETEXPLORER
        caps['initialBrowserUrl'] =url
        driver = webdriver.Ie(executable_path=exc_path_ie,capabilities=caps)
        driver.maximize_window()
        #driver.get(url)
        sleep(3)
        driver.find_element_by_xpath("//input[@id='Ecom_User_ID']").send_keys(userid)    
        ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@id='Ecom_Password']")))).click().send_keys(pswd).send_keys(Keys.RETURN).perform()
        sleep(5)
        if "svp" in url:
            WebDriverWait(driver,20).until(EC.presence_of_element_located((By.XPATH,"//a[@id='ui-id-97']"))).click()
        else:
            WebDriverWait(driver,20).until(EC.presence_of_element_located((By.XPATH,"//a[@id='ui-id-103']"))).click()
        ActionChains(driver).move_to_element(WebDriverWait(driver,15).until(EC.element_to_be_clickable((By.XPATH,"//input[@type='text' and @name = 's_1_1_2_0']")))).click().send_keys(cust).send_keys(Keys.RETURN).perform()
        sleep(3)
        ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH,"//table[@id='s_1_l']//td[@id='1_s_1_l_Last_Name']//a[@name = 'Last Name']")))).click().perform()
        sleep(3)
        cursor = WebDriverWait(driver,20).until(EC.element_to_be_clickable((By.XPATH,"//a[@id='s_13_1_1_0_mb']")))
        driver.execute_script("arguments[0].scrollIntoView()",cursor)
        ActionChains(driver).move_to_element(cursor).click().perform()
        select = Select(WebDriverWait(driver,60).until(EC.element_to_be_clickable((By.XPATH,"//select[@id='j_s_vctrl_div_tabScreen']"))))
        select.select_by_visible_text('Products')
        elements = WebDriverWait(driver,20).until(EC.presence_of_all_elements_located((By.XPATH,"//table[@id='s_3_l']/tbody/tr")))
        for i,elem in enumerate(elements):
            if 'BT SUPER FOR LIFE - SAVINGS' in elem.text or 'BT SFL' in elem.text:
                WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH,"//table[@id='s_3_l']//tr[@id='"+str(i)+"']"))).click()
                ActionChainsbtstg(driver).move_to_element(WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH,"//span[contains(text(),'Service')]")))).click().perform()
                break
        try:
            WebDriverWait(driver, 5).until(EC.alert_is_present())

            alert = driver.switch_to.alert
            alert.accept()
            return 'pass'
        except Exception as e:
            driver.switch_to.frame('symbUrlIFrame0')
            sleep(3)
            cursor = WebDriverWait(driver,30).until(EC.presence_of_element_located((By.XPATH,'//div[@class="sfl-section-heading"]| //article[@class="unauthenticated-text"]')))
            if cursor.text == "Sorry, The page you have requested is only available after you have signed in to Internet Banking.":
                    raise Exception()
            return 'pass'
    except Exception as e:
        return 'fail'
        
    finally:
        try:
            driver.close()
            driver.quit()
        except:
            pass
            

    
def ingress_stg():
    res=[]
    stg_input=[]
    ind =['btfsl','gcis','env']
    env = ['','','SIT1','SIT2','SIT3','UAT1','SVP']
    for col in range(2,7):
        temp=[]
        for row in (5,6):
            temp += [sheet.cell_value(row,col).strip()]
            
        stg_input +=[ dict(zip(ind,temp+[env[col]]))]


    col = [[3,4,5,6,7],[8,9,10,11,12,13],[14,15],[16,17,18,19],[20,21,22]]
    for i,data in enumerate(stg_input):
        print("stg "+ data["env"])
        res+=[btstg(data["gcis"],data["btfsl"],data["env"])]
        ud.update([15],col[i],res[-1],"BT")
    return res
    
def ingress_firefly():
    res=[]
    firefly_input=[]
    ind =['btfsl','gcis','env']
    env = ['','','SIT1_SAML_FF','SIT2_SAML_FF','SIT3_SAML_FF','UAT1_SAML_FF']
    for col in range(2,6):
    #for col in range(2,4):
        temp=[]
        for row in (8,9):
            temp += [sheet.cell_value(row,col).strip()]   
        firefly_input +=[ dict(zip(ind,temp+[env[col]]))]
    col = [[3,4,5,6,7],[8,9,10,11,12,13],[14,15],[16,17,18,19]]
    #col = [[3,4,5,6,7],[8,9,10,11,12,13]]
    for i,data in enumerate(firefly_input):
        print("FireFly "+ data["env"])
        res+=[btstg(data["gcis"],data["btfsl"],data["env"])]
        ud.update([16],col[i],res[-1],"BT")
    return res

def ingress_wbc():
    res=[]
    wbc_input=[]
    ind =['can','env']
    env = ['','','SIT1','SIT2','SIT3','UAT1','SVP']
    for col in range(2,7):
    #for col in range(2,4):
        wbc_input +=[ dict(zip(ind,[sheet.cell_value(11,col).strip(),env[col]]))]
    col = [[3,4,5,6,7],[8,9,10,11,12,13],[14,15],[16,17,18,19],[20,21,22]]
    #col = [[3,4,5,6,7],[8,9,10,11,12,13]]  
    for i,data in enumerate(wbc_input):
        print("WBC "+ data["env"])
        res+=[btwbc(data["can"],data["env"])]
        ud.update([17,18],col[i],res[-1],"BT")
    return res

def ingress_rb():
    res=[]
    rb_input=[]
    ind =['can','env']
    env = ['','','SIT1_NTLM_RB','SIT2_NTLM_RB','SIT3_NTLM_RB','UAT1_NTLM_RB']
    for col in range(2,6):
    #for col in range(2,4):
        rb_input +=[ dict(zip(ind,[sheet.cell_value(13,col).strip(),env[col]]))]
    col = [[3,4,5,6,7],[8,9,10,11,12,13],[14,15],[16,17,18,19]]
    #col = [[3,4,5,6,7],[8,9,10,11,12,13]]    
    for i,data in enumerate(rb_input):
        print("Rb "+ data["env"])
        res+=[btwbc(data["can"],data["env"])]
        ud.update([19],col[i],res[-1],"BT")
    return res

def compass_d():
    res=[]
    compass_input=[]
    ind =['can','scn','itp','link']
    for col in (2,3):
        temp=[]
        for row in (15,16,17,18):
            temp +=[sheet.cell_value(row,col).strip()]
        compass_input +=[ dict(zip(ind,temp))]
    col = [[3,4,5,6,7],[8,9,10,11,12,13]]
    for i,data in enumerate(compass_input):
        print("Compass Webpage: "+ data["link"])
        res+=[compass_desktop(data["can"],data["scn"],data["itp"],data["link"])]
        ud.update([20,21,22,23,24,25,26,27],col[i],res[-1],"BT")
    return res
    
def compass_m():
    res=[]
    compass_input=[]
    ind =['can','scn','itp','link']
    for col in (2,3):
        temp=[]
        for row in (20,21,22,23):
            temp +=[sheet.cell_value(row,col).strip()]
        compass_input +=[ dict(zip(ind,temp))]
    col = [[3,4,5,6,7],[8,9,10,11,12,13]]      
    for i,data in enumerate(compass_input):
        print("Compass Webpage: "+ data["link"])
        res+=[compass_mobile(data["can"],data["scn"],data["itp"],data["link"])]
        ud.update([28,29,30,31],col[i],res[-1],"BT")
    return res



def westpac_live():
    res=[]
    wlive_input=[]
    ind =['can','pswd','link']
    for col in (2,3,4,5):
    #for col in range(2,4):
        temp=[]
        for row in (30,31,32):
            temp +=[sheet.cell_value(row,col).strip()]
        wlive_input +=[ dict(zip(ind,temp))]
    col = [[3,4,5,6,7],[8,9,10,11,12,13],[14,15],[16,17,18,19]]
    #col = [[3,4,5,6,7],[8,9,10,11,12,13]]
    for i,data in enumerate(wlive_input):
        print("Westpac Live Webpage: "+ data["link"])
        res+=[wlive(data["can"],data["pswd"],data["link"])]
        ud.update([33,34,35,36,37,38],col[i],res[-1],"BT")
    return res
    

def relationship_builder():
    res=[]
    rb_input=[]
    ind =['customer','userid','pswd','link']
    for col in (2,3,5):
    #for col in range(2,4):
        temp=[]
        for row in (34,35,36,37):
            temp +=[sheet.cell_value(row,col).strip()]
        rb_input +=[ dict(zip(ind,temp))]
    col = [[3,4,5,6,7],[8,9,10,11,12,13],[16,17,18,19]]
    #col = [[3,4,5,6,7],[16,17,18,19]]
    for i,data in enumerate(rb_input):
        print("RelationShip Builder Webpage: "+ data["link"])
        res+=[rb(data['customer'],data["userid"],data["pswd"],data["link"])]
        ud.update([39],col[i],res[-1],"BT")
    
    return res

def apply_proxy_to_ie():
    driver = webdriver.Ie(executable_path=exc_path_ie)
    driver.get("http://dwwas0004.btfin.com:9081/ingress/access")
    driver.quit()

def sendmail():
    today_date = date.today()    
    todays_date = str(today_date.day)+'th '+today_date.strftime('%B')+' '+str(today_date.year)+', '+today_date.strftime('%A')
    
    xl = EnsureDispatch('Excel.Application')
    wb = xl.Workbooks.Open(r'P:\imran-TEMS\BTFSL\BT Report.xlsx')
    ws = wb.Worksheets('Test Report')
    ws.Cells(8,'R').Value = str(today_date.day)+'th '+today_date.strftime('%B')+' '+str(today_date.year)+', '+today_date.strftime('%A')
    wb.PublishObjects.Add(SourceType=constants.xlSourceRange,Filename=r'P:\imran-TEMS\BTFSL\BT Report.html',Sheet='Test Report',Source='$A$1:$AC$70',HtmlType=constants.xlHtmlStatic, DivID='xxx1')
    wb.PublishObjects(1).Publish(True)
    wb.Close(True)
    xl.Application.Quit()
    
    
    body_content = open(r'P:\imran-TEMS\BTFSL\BT Report.html').read()
    path_of_img = {1:'P:\imran-TEMS\dummy\Tems.png',0:'P:\imran-TEMS\dummy\westpaclogo.jpg'}
    count=0
    num = 0
     
    while count<2 and num<50:
        pat = "BT%20Report_files/xxx1_image00"+str(num)+".png"
        
        match = re.search(pat,body_content)
        body_content = re.sub(pat,path_of_img[count],body_content)
        num+=1
        if match is None:
            continue
        count+=1

    outlook = Dispatch('outlook.application')
    user_id=getpass.getuser()

    mail = outlook.CreateItem(0)
    #mail.To = "TEMS Service Desk <temsservicedesk@westpac.com.au>"
    #mail.BCC = "L096535;L099582;L098864;L092160;L064403;L100513;Sveum, Mikkal <Mikkal.Sveum@BTFinancialgroup.com>; Baird, Russell <russellbaird@westpac.com.au>; Brinkman, Derek <dbrinkman@westpac.com.au>;L062449;L066667;"
    mail.To = "L111185"
    mail.BCC = "L092160"
    mail.Subject = "BTSFL Test Environment Status(Automated) "+ str(today_date.day)+'th '+today_date.strftime('%B')+" "+str(today_date.year)
    mail.HTMLbody  = body_content
    mail.send
    os.remove("P:\imran-TEMS\BTFSL\BT Report.html")
    
def main():
    print("Run started at : "+str(datetime.now()))
    logging.basicConfig(filename=r'P:\imran-TEMS\BTFSL\BTFSLLogs.log',level=logging.INFO,filemode='w',format='%(asctime)s %(message)s',datefmt='%d/%m/%Y %I:%M:%S %p')
    logging.info('BTFSL for STG Customer')
    os.system("taskkill /f /im EXCEL.EXE")
    
    os.system('''echo y|reg add "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings" /v AutoConfigURL /d "http://proxycfg.btfin.com/wpad-acc.dat"''')
    #apply_proxy_to_ie()
    ud.start_xl()
    stg= ingress_stg()
    print(stg)
    
    ud.kill_xl()
    
    
    os.system("taskkill /f /im EXCEL.EXE")

    
    logging.info("Script Completed")
    print("Run completed at : "+str(datetime.now()))
    
    
if __name__=="__main__":
    main()
