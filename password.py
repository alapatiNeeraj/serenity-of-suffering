import tkinter as tk
import tkinter.scrolledtext as tkscrolled
from tkinter import StringVar, IntVar
from selenium import webdriver
from time import sleep
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
import sys,re,xlsxwriter
from xlrd import open_workbook
import pandas as pd


def create_frames():
    global lid,userEnteredRoles,rolesRemoved
    #frame 1
    tk.Label(f1,text="Login ID",width=20).grid(row=0,column=0,columnspan = 2,sticky=tk.W,pady=10)
    tk.Entry(f1,textvariable=user_lid,width=30).grid(row=0,column=2,columnspan=2,sticky=tk.E,padx=5,pady=10)
    tk.Label(f1,text="Password",width=20).grid(row=1,column=0,columnspan = 2,sticky=tk.W,pady=10)
    tk.Entry(f1,show='*',textvariable = user_password,width=30).grid(row=1,column=2,columnspan = 2,sticky=tk.E,padx=5,pady=10)
    tk.Button(f1, text="Login",command=login).grid(row=2,column=0,columnspan = 2,ipadx=30,padx=20,pady=10,sticky=tk.W)
    tk.Button(f1, text="Quit",command=destroy).grid(row=2,column=2,columnspan = 2,ipadx=30,padx=20,pady=10,sticky=tk.E)
    tk.Radiobutton(f1, text="SIT",variable=url,value="https://sso-sit.intranet.westpac.com.au/helpdesk/ibm/console/").grid(row=3,column=1)
    tk.Radiobutton(f1, text="UAT",variable=url,value="https://sso-svp.intranet.westpac.com.au/helpdesk/ibm/console/").grid(row=3,column=2)

    msgf1.set("Note:If both username and password field is left empty then default ID(L096535) will be used to login")
    tk.Label(f1,textvariable=msgf1,wraplength=340,justify=tk.LEFT).grid(row=4,columnspan=4,padx=10,pady=(0,10))

    #Frame 2
    tk.Radiobutton(f2, text="Password Reset",variable=val,value=1).grid(row=0,column=0,sticky=tk.W,pady = 10)
    tk.Radiobutton(f2, text="Access Request",variable=val,value=2).grid(row=0,column=1,sticky=tk.E,pady=10)
    tk.Radiobutton(f2, text="Access Request and Password Reset",variable=val,value=3).grid(row=1,column=0,sticky=tk.W)
    tk.Radiobutton(f2, text="Roles removal",variable=val,value=4).grid(row=1,column=1,sticky=tk.E)
    msgf2.set("Please enter Lid or LId's separated by either space,comma(,) or enter.")
    tk.Label(f2,textvariable=msgf2,wraplength=350,justify=tk.LEFT).grid(row=2,columnspan=2, sticky="NSEW")
    lid=tkscrolled.ScrolledText(f2,width=45,height=4)
    lid.grid(row=3,columnspan=2,sticky="NSEW",padx=5)
    tk.Button(f2, text="Continue",command=cont).grid(row=4,column=0,columnspan = 2,ipadx=30,padx=20,pady=10,sticky="NSEW")

    #Frame 3
    tk.Label(f3, text = "Do you have any model Id or do you want to enter roles?",wraplength=350,justify=tk.LEFT).grid(row=0,columnspan=2, padx = (10,0),sticky=tk.W)
    tk.Radiobutton(f3,text="Model Id",variable=mod,value=1,command = model).grid(row=1,column=0,sticky = tk.W,padx=(10,0))
    tk.Radiobutton(f3,text="Provide specific roles",variable=mod,value=2,command = model).grid(row=1,column=1,sticky = tk.E,padx=(0,10))
    tk.Label(f3,textvariable=msgf3,wraplength=340,justify=tk.LEFT).grid(row=2,columnspan=2, padx=(10,0),sticky="NSEW")
    tk.Button(f3, text="Continue",command=access).grid(row=5,column=1,ipadx=30,padx=20,pady=10,sticky=tk.E)
    tk.Button(f3, text="Back",command=f2.tkraise).grid(row=5,column=0,ipadx=30,padx=20,pady=10,sticky=tk.W)

    #frame 4
    tk.Label(f4,text="Do you have specific roles to be removed or you want to remove roles of specific applications?",wraplength=340,justify=tk.LEFT).grid(row=0,columnspan=2, padx=10,sticky=tk.W)
    tk.Radiobutton(f4,text="Remove by app name",variable=rem,value=1,command = select_for_removal).grid(row=1,column=0,sticky = tk.W,padx=(10,0))
    tk.Radiobutton(f4,text="Remove by roles name",variable=rem,value=2,command = select_for_removal).grid(row=1,column=1,sticky = tk.E,padx=(0,10))
    tk.Label(f4,textvariable=msgf4,wraplength=340,justify=tk.LEFT).grid(row=2,columnspan=2, padx=(10,0),sticky="NSEW")
    tk.Button(f4, text="Back",command=f2.tkraise).grid(row=5,column=0,ipadx=30,padx=20,pady=10,sticky=tk.W)
    tk.Button(f4, text="Continue",command=roles_removal).grid(row=5,column=1,ipadx=30,padx=20,pady=10,sticky=tk.E)
    
    
    #frame 5
    tk.Label(f5,text="Please provide password in below text box and press password reset button.",wraplength=340,justify=tk.LEFT).grid(row=0,column=0,columnspan=2, padx=(10,0),sticky="NSEW")
    tk.Label(f5,text="Note: All users password will be reset to same password.",wraplength=340,justify=tk.LEFT).grid(row=1,column=0,columnspan=2, padx=(10,0),sticky="NSEW")
    tk.Label(f5,text="Password : ").grid(row=3,column=0,padx=(20,0),sticky=tk.W)
    tk.Entry(f5,textvariable=pass_change,width=25).grid(row=3,column=1,padx=10,sticky=tk.W)
    tk.Button(f5, text="Password Reset",command=pass_reset).grid(row=4,column=0,ipadx=10,padx=20,pady=10,sticky=tk.W)
    tk.Button(f5, text="Logout",command=initial_screen).grid(row=4,column=1,ipadx=30,padx=20,pady=10,sticky=tk.W)    
    tk.Label(f5,textvariable=msgf5,wraplength=340,justify=tk.LEFT).grid(row=5,column=0,columnspan=2, padx=(10,0),sticky="NSEW")

    #frame 6
    tk.Label(f6,text="Do you want to update access request in portal?",wraplength=340,justify=tk.LEFT).grid(row=0,column=0,columnspan=2, padx=(10,0),sticky="NSEW")
    tk.Radiobutton(f6,text="Yes",variable=port,value=1,command = portal).grid(row=1,column=0,sticky = tk.W,padx=(10,0))
    tk.Radiobutton(f6,text="No",variable=port,value=2,command = portal).grid(row=1,column=1,sticky = tk.E,padx=(0,10))
    tk.Label(f6,textvariable=msgf6,wraplength=340,justify=tk.LEFT).grid(row=4,column=0,columnspan=2, padx=(10,0),sticky="NSEW")
    
    
def destroy():
    driver.quit()
    root.destroy()



def login():
    global driver
    if url.get()=="1":
        msgf1.set("Please select environment")
    else:
        if (user_lid.get() != "" and user_password.get()=="") or (user_lid.get() == "" and user_password.get() !=""):
            msgf1.set("Username or Password is empty. Please check and provide all details.")
        else:
            if user_lid.get() == "" and user_password.get()=="":
                user_lid.set("L096535")
                user_password.set("Westpac19")
                wb=open_workbook(r'P:\imran-TEMS\password\default.xlsx')
                sheet=wb.sheet_by_index(0)
                user_lid.set(sheet.cell_value(1,0))
                user_password.set(sheet.cell_value(1,1))
                msgf1.set("Logging in with default username and ID")
            else:
                msgf1.set("Logging In. Please Wait...")
            chrome_options=Options()
            driver.get(url.get())
            WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@type='TEXT']"))).send_keys(user_lid.get())
            WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@type='PASSWORD']"))).send_keys(user_password.get())
            WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@type='submit']"))).click()
            try:
                incorrect = WebDriverWait(driver,3).until(EC.presence_of_element_located((By.XPATH,"//html//table[3]/tbody[1]/tr[1]/td[1]")))
                if "HPDIA0200W Authentication failed. You have used an invalid user name, password or client certificate." in incorrect.text:  
                    msgf1.set("HPDIA0200W Authentication failed. You have used an invalid user name, password or client certificate.")
                    user_lid.set("")
                    user_password.set("")
                    f1.tkraise()
            except TimeoutException:
                try:
                    WebDriverWait(driver,3).until(EC.presence_of_element_located((By.XPATH,"//h1[contains(text(),'Not Found')]")))
                    driver.get(url.get())
                except:
                    pass
                try:
                   WebDriverWait(driver,3).until(EC.presence_of_element_located((By.XPATH,"//input[@type='submit']"))).click()
                except:
                    pass
                f2.tkraise()


def pass_reset():
    global driver,lid,userEnteredRoles
    if len(pass_change.get())==0:
        msgf5.set("Please provide password in the box above before pressing Password Reset button")
    else:
        driver.switch_to.default_content()
        cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//frame[@name='navigation']")))
        driver.switch_to.frame(cursor)
        ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//span[contains(text(),'IBM Security Access Manager')]")))).click().perform()
        ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//span[contains(text(),'Web Portal Manager')]")))).click().perform()
        ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"(//span[contains(text(),'Users')])[last()]")))).click().perform()
        id_list = re.split(",| |\n",lid.get("1.0",'end-1c'))
        id_list = [i.strip() for i in id_list]
        id_list = list(filter(None, id_list))
        try:
            for ids in id_list:
                driver.switch_to.default_content()
                cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//frame[@name='navigation']")))
                driver.switch_to.frame(cursor)
                ActionChains(driver).move_to_element(WebDriverWait(driver,3).until(EC.presence_of_element_located((By.XPATH,"//a[contains(text(),'Search Users')]")))).click().perform()
                driver.switch_to.default_content()
                driver.switch_to.frame(driver.find_element_by_xpath("//frame[@title='Content frame']"))
                sleep(3)
                driver.switch_to.frame('IFRAME_DEMO')
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@id='filter']")))).double_click().send_keys(ids).send_keys(Keys.RETURN).perform()
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.ID,"pwd")))).click().send_keys(pass_change.get()).perform()
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.ID,"vpwd")))).click().send_keys(pass_change.get()).send_keys(Keys.RETURN).perform()
                sleep(1)
                WebDriverWait(driver,10).until(EC.presence_of_element_located((By.ID,"pwd"))).click()
            f6.tkraise()    
        except:
            try:
                WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//p[contains(text(),'HPDIA0300W   Password rejected due to policy violation. (0x1321212c)')]")))
                msgf5.set("Password that you have provided got rejected due to policy violation. Please provide another password.")
                driver.switch_to.default_content()
                cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//frame[@name='navigation']")))
                driver.switch_to.frame(cursor)
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"(//span[contains(text(),'Users')])[last()]")))).click().perform()
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//span[contains(text(),'Web Portal Manager')]")))).click().perform()
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//span[contains(text(),'IBM Security Access Manager')]")))).click().perform()
                driver.switch_to.default_content()
                pass_change.set("")
                f5.tkraise()
            except:
                WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//p[contains(text(),'HPDMG0754W   The entry was not found. If a user or group is being created, ensure that the Distinguished Name (DN) specified has the correct syntax and is valid. (0x14c012f2)')]")))
                driver.switch_to.default_content()
                cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//frame[@name='navigation']")))
                driver.switch_to.frame(cursor)
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"(//span[contains(text(),'Users')])[last()]")))).click().perform()
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//span[contains(text(),'Web Portal Manager')]")))).click().perform()
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//span[contains(text(),'IBM Security Access Manager')]")))).click().perform()
                msgf2.set("The User Lid "+ids+" that you have provided deos not exist. Please provide a valid user LId.")
                f2.tkraise()

                
   


def access():
    global driver,lid,userEnteredRoles
    if mod.get()== 0:
        msgf3.set("Please select one of the two option above.")
        f3.tkraise()
    elif ((model_id.get() == "" or app_name.get() == "") and mod.get()==1) or (userEnteredRoles.get("1.0","end-1c")=="" and mod.get()==2):
        msgf3.set("Please provide all the details before pressing continue.")
        f3.tkraise()
    else:
        driver.switch_to.default_content()
        cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//frame[@name='navigation']")))
        driver.switch_to.frame(cursor)
        ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//span[contains(text(),'IBM Security Access Manager')]")))).click().perform()
        ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//span[contains(text(),'Web Portal Manager')]")))).click().perform()
        ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"(//span[contains(text(),'Users')])[last()]")))).click().perform()
        if mod.get()==1:
            ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//a[contains(text(),'Search Users')]")))).click().perform()
            driver.switch_to.default_content()
            driver.switch_to.frame(driver.find_element_by_xpath("//frame[@title='Content frame']"))
            sleep(3)
            driver.switch_to.frame('IFRAME_DEMO')
            ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@id='filter']")))).double_click().send_keys(model_id.get().strip()).send_keys(Keys.RETURN).perform()
            try:
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//a[contains(text(),'Groups')]")))).click().perform()
            except:
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//p[contains(text(),'HPDMG0754W   The entry was not found. If a user or group is being created, ensure that the Distinguished Name (DN) specified has the correct syntax and is valid. (0x14c012f2)')]")))).click().perform()
                driver.switch_to.default_content()
                cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//frame[@name='navigation']")))
                driver.switch_to.frame(cursor)
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"(//span[contains(text(),'Users')])[last()]")))).click().perform()
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//span[contains(text(),'Web Portal Manager')]")))).click().perform()
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//span[contains(text(),'IBM Security Access Manager')]")))).click().perform()
                driver.switch_to.default_content()
                model_id.set("")
                msgf3.set("HPDMG0754W   The entry was not found. Model ID provided is incorrect please provide correct model ID")
                f3.tkraise()
            
            cursor = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//table[@class='table-border']")))
            model_roles=cursor.text.split("\n")
            model_roles.remove('  Select Group Name')
            model_roles = [i.strip() for i in model_roles]
            model_roles = list(filter(None, model_roles))
            app = app_name.get().lower().split(",")
            if 'all' in app:
                app = list(set([re.split('-|_',i)[0] for i in model_roles]))
                
            id_list = re.split(",| |\n",lid.get("1.0",'end-1c'))
            id_list = [i.strip() for i in id_list]
            id_list = list(filter(None, id_list))
                
        elif mod.get()==2:
            id_list = re.split(",| |\n",lid.get("1.0",'end-1c'))
            id_list = [i.strip() for i in id_list]
            id_list = list(filter(None, id_list))
            
            model_roles = re.split(" ,| |\n",userEnteredRoles.get("1.0",'end-1c'))
            model_roles = [i.strip() for i in model_roles]
            model_roles = list(filter(None, model_roles))
            
            app = list(set([re.split('-|_',i)[0] for i in model_roles]))
            
        for ids in id_list:
            driver.switch_to.default_content()
            cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//frame[@name='navigation']")))
            driver.switch_to.frame(cursor)
            ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//a[contains(text(),'Search Users')]")))).click().perform()
            driver.switch_to.default_content()
            driver.switch_to.frame(driver.find_element_by_xpath("//frame[@title='Content frame']"))
            sleep(3)
            driver.switch_to.frame('IFRAME_DEMO')

            ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@id='filter']")))).double_click().send_keys(ids).send_keys(Keys.RETURN).perform()
            try:
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//a[contains(text(),'Groups')]")))).click().perform()
            except:
                req_roles={}                
                for i in app:
                    req_roles[i]=[role for role in model_roles if role.startswith(i)]
                    
                WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//p[contains(text(),'HPDMG0754W   The entry was not found. If a user or group is being created, ensure that the Distinguished Name (DN) specified has the correct syntax and is valid. (0x14c012f2)')]")))
                window_before = driver.window_handles[0]
                driver.execute_script("window.open('https://intranet.westpacgroup.com.au/wbg/home', 'tab2');")
                driver.switch_to_window(driver.window_handles[1])
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@id='newPDQuery']")))).click().send_keys(ids).send_keys(Keys.RETURN).perform()
                sleep(5)
                try:
                    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//div[@id='no-results']")))
                    driver.close()
                    msgf2.set(ids+" is not a valid Id. Please recheck the id ")
                    driver.switch_to_window(window_before)
                    driver.switch_to.default_content()
                    cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//frame[@name='navigation']")))
                    driver.switch_to.frame(cursor)
                    ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"(//span[contains(text(),'Users')])[last()]")))).click().perform()
                    ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//span[contains(text(),'Web Portal Manager')]")))).click().perform()
                    ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//span[contains(text(),'IBM Security Access Manager')]")))).click().perform()
                    f2.tkraise()
                    sys.exit(0)
                except:
                    pass
                name = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//div[@class='name ng-binding']"))).text
                driver.close()
                driver.switch_to_window(window_before)
                driver.switch_to.default_content()  
                cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//frame[@name='navigation']")))
                driver.switch_to.frame(cursor)
                
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//a[contains(text(),'Create User')]")))).click().perform()
                driver.switch_to.default_content()
                driver.switch_to.frame(driver.find_element_by_xpath("//frame[@title='Content frame']"))
                sleep(3)
                driver.switch_to.frame('IFRAME_DEMO')
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.ID,"id")))).click().send_keys(ids).perform()
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.ID,"ldapcn")))).click().send_keys(name.split(" ")[0]).perform()
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.ID,"ldapsn")))).click().send_keys(name.split(" ")[1]).perform()
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.ID,"pwd")))).click().send_keys("initpass1").perform()
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.ID,"vpwd")))).click().send_keys("initpass1").perform()
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@type='submit' and @value= 'Create']")))).click().perform()
                sleep(5)
                driver.switch_to.default_content()
                cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//frame[@name='navigation']")))
                driver.switch_to.frame(cursor)
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//a[contains(text(),'Search Users')]")))).click().perform()
                driver.switch_to.default_content()
                driver.switch_to.frame(driver.find_element_by_xpath("//frame[@title='Content frame']"))
                sleep(3)
                driver.switch_to.frame('IFRAME_DEMO')
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@id='filter']")))).double_click().send_keys(ids).send_keys(Keys.RETURN).perform()
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//a[contains(text(),'Groups')]")))).click().perform()                
                WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@type='button' and @value='Add...']"))).click()
     
                for i in req_roles:
                    if len(req_roles[i])>0:
                        ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@id='max']")))).double_click().send_keys('1000').perform()
                        ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@id='filter']")))).double_click().send_keys(Keys.DELETE).double_click().send_keys(i+"*").send_keys(Keys.RETURN).perform()
                        action = ActionChains(driver).key_down(Keys.LEFT_CONTROL)
                        try:
                            
                            for role in req_roles[i]:
                                WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//option[contains(text(),'"+role+"')]"))).click()
                            action.click(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="grouplist"]/option[1]'))))
                            action.key_up(Keys.LEFT_CONTROL).send_keys(Keys.RETURN).perform()
                            WebDriverWait(driver,20).until(EC.presence_of_element_located((By.XPATH,"//p[contains(text(),'The user was added to the groups successfully')]")))
                        except:
                            cursor = WebDriverWait(driver,5).until(EC.presence_of_element_located((By.XPATH,"//td[contains(text(),'No groups matched the search criteria')] | //table[@class = 'msgtable']")))
                            
                            if cursor.text == "No groups matched the search criteria":
                                print("The role : "+i+" was not found")
                                continue
                            elif "The specified group is a dynamic group and cannot be modified" in cursor.text:
                                WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@type='button' and @value ='Back']"))).click()
                                print("The role : "+i+" was not found as it is a dynamic group")
                                continue
                WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@type='button' and @value= 'Done']"))).click()

                continue
                        

            cursor = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//table[@class='table-border']")))            
            id_roles=cursor.text.split("\n")
            roles_to_be_added = list(set(model_roles)-set(id_roles))
            req_roles={}

            for i in app:
                req_roles[i]=[role for role in roles_to_be_added if role.startswith(i)]

            WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@type='button' and @value='Add...']"))).click()
            action = ActionChains(driver).key_down(Keys.LEFT_CONTROL)
            for i in req_roles:
                if len(req_roles[i])>0:
                    ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@id='max']")))).double_click().send_keys('1000').perform()
                    ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@id='filter']")))).double_click().send_keys(Keys.DELETE).double_click().send_keys(i+"*").send_keys(Keys.RETURN).perform()
                    action = ActionChains(driver).key_down(Keys.LEFT_CONTROL)
                    try:
                        
                        for role in req_roles[i]:
                            #print(role)
                            WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//option[contains(text(),'"+role+"')]"))).click()
                        action.click(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="grouplist"]/option[1]'))))
                        action.key_up(Keys.LEFT_CONTROL).send_keys(Keys.RETURN).perform()
                        WebDriverWait(driver,20).until(EC.presence_of_element_located((By.XPATH,"//p[contains(text(),'The user was added to the groups successfully')]")))
                    except:
                        cursor = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//td[contains(text(),'No groups matched the search criteria')] | //table[@class = 'msgtable']")))
                            
                        if cursor.text == "No groups matched the search criteria":
                            print("The role : "+i+" was not found")
                            continue
                        elif "The specified group is a dynamic group and cannot be modified" in cursor.text:
                            WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@type='button' and @value ='Back']"))).click()
                            print("The role : "+i+" was not found as it is a dynamic group")
                            continue            
                    
            WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@type='button' and @value= 'Done']"))).click()
        
            sleep(2)
        driver.switch_to.default_content()
        cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//frame[@name='navigation']")))
        driver.switch_to.frame(cursor)
        ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"(//span[contains(text(),'Users')])[last()]")))).click().perform()
        ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//span[contains(text(),'Web Portal Manager')]")))).click().perform()
        ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//span[contains(text(),'IBM Security Access Manager')]")))).click().perform()
        is_password()

        
def roles_removal():
    global driver,lid,userEnteredRoles
    if rem.get()== 0:
        msgf4.set("Please select one of the two option above.")
        f4.tkraise()
    elif ((app_name_remove.get() == "") and rem.get()==1) or (rolesRemoved.get("1.0","end-1c")=="" and rem.get()==2):
        msgf4.set("Please provide all the details before pressing continue.")
        f4.tkraise()
    else:
        id_list = re.split(",| |\n",lid.get("1.0",'end-1c'))
        id_list = [i.strip() for i in id_list]
        id_list = list(filter(None, id_list))
        if rem.get()==1:
            app = app_name_remove.get().lower().split(",")
        elif rem.get()==2:
            userRoles = re.split(" ,| |\n",rolesRemoved.get('1.0','end-1c'))
            userRoles = [i.strip() for i in userRoles]
            userRoles = list(filter(None, userRoles))

        driver.switch_to.default_content()
        cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//frame[@name='navigation']")))
        driver.switch_to.frame(cursor)
        ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//span[contains(text(),'IBM Security Access Manager')]")))).click().perform()
        ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//span[contains(text(),'Web Portal Manager')]")))).click().perform()
        ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"(//span[contains(text(),'Users')])[last()]")))).click().perform()
        
        for ids in id_list:
            driver.switch_to.default_content()
            cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//frame[@name='navigation']")))
            driver.switch_to.frame(cursor)
            ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//a[contains(text(),'Search Users')]")))).click().perform()
            driver.switch_to.default_content()
            driver.switch_to.frame(driver.find_element_by_xpath("//frame[@title='Content frame']"))
            sleep(3)
            driver.switch_to.frame('IFRAME_DEMO')
            ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@id='filter']")))).double_click().send_keys(ids).send_keys(Keys.RETURN).perform()
            
            try:                
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//a[contains(text(),'Groups')]")))).click().perform()
                if rem.get()==1:
                    cursor = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//table[@class='table-border']")))            
                    id_roles=cursor.text.split("\n")
                    userRoles = [i for i in id_roles if re.split('-|_',i)[0] in app]

                    
                for role in userRoles:
                    try:
                        cursor=WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@type='checkbox' and @value='"+role+"']")))
                        cursor.click()
                    except:
                        pass
                try:
                    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@type='submit' and @value='Remove']"))).click()
                    cursor=WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@type='button' and @value='Remove']")))
                    driver.execute_script("arguments[0].scrollIntoView()", cursor)
                    cursor.click() 
                    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//a[contains(text(),'Groups')]")))
                except:
                    pass
            except:
                WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//p[contains(text(),'HPDMG0754W   The entry was not found. If a user or group is being created, ensure that the Distinguished Name (DN) specified has the correct syntax and is valid. (0x14c012f2)')]")))
                driver.switch_to.default_content()
                cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//frame[@name='navigation']")))
                driver.switch_to.frame(cursor)
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"(//span[contains(text(),'Users')])[last()]")))).click().perform()
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//span[contains(text(),'Web Portal Manager')]")))).click().perform()
                ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//span[contains(text(),'IBM Security Access Manager')]")))).click().perform()
                msgf2.set("The User Lid "+ids+" that you have provided deos not exist. Please provide a valid user LId.")
                f2.tkraise()
                sys.exit(0)
                

        initial_screen()

def access_request_portal():
    if port.get() == 1:
        if accessNo.get() == "" or req.get() ==0:
            msgf6.set("Please provide all the details before pressing continue.")
            f6.tkraise()
        else:
            app = app_name.get().upper().split(",")
            id_list = re.split(",| |\n",lid.get("1.0",'end-1c'))
            id_list = [i.strip() for i in id_list]
            id_list = list(filter(None, id_list))
    
            Environment = ["SIT/ESIT" if url.get()=="https://sso-sit.intranet.westpac.com.au/helpdesk/ibm/console/" else "UAT/EUAT"]*len(id_list)
            SalaryId = id_list
            Password = [pass_change.get()]*len(id_list)
            try:
                Application = [app[0]]*len(id_list)
            except:
                Application = [" "]*len(id_list)
            Comments = ["Access has been provided" if req.get()==2 else "Password has been reset"]*len(id_list)
            
            data = pd.DataFrame({'Salary Id':SalaryId,'Password':Password,'Application':Application,'Environment': Environment,'Comments':Comments})
            column_order = ['Salary Id','Password','Application','Environment','Comments']
            data = data[column_order]
            path = 'H:\\Documents\\' +accessNo.get()+ '.xlsx'
            writer= pd.ExcelWriter(path,engine='xlsxwriter')
            data.to_excel(writer,sheet_name='details',index=False)
            workbook  = writer.book
            worksheet1 = writer.sheets['details']
            header_format = workbook.add_format({
                'bold': True,
                'align':'center',
                'valign': 'vcenter',
                'text_wrap': True,
                'fg_color': '#ffffff',
                'border': 1})


            for col_num, value in enumerate(data.columns.values):
                worksheet1.write(0, col_num, value, header_format)
            format1 =workbook.add_format({'text_wrap': True,'align':'center','valign': 'vcenter','border': 1})
            worksheet1.set_column('A:E', 15,format1)
            worksheet1.set_column('A:A', 15)
            worksheet1.set_column('B:B', 15)
            worksheet1.set_column('C:C', 15)
            worksheet1.set_column('D:D', 15)
            worksheet1.set_column('E:E', 27)
            
            worksheet1.set_row(0,16)

            writer.save()
            writer.close()
            window_before = driver.window_handles[0]
            driver.execute_script("window.open('http://lisa-wkstn-tcp.dev.srv.westpac.com.au:8083/admin/ServiceDeskTeam.aspx', 'tab2');")
            driver.switch_to_window(driver.window_handles[1])
            try:
                WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@id='username']"))).send_keys("sd")
                WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@id='password']"))).send_keys("sd2016")
                WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@id='loginW']"))).click()
            except:
                pass
            ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH,"//span[@class='sr-only']")))).click().perform()
            ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH,"//span[contains(text(),'Reports')]")))).click().perform()
            ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH,"//a[@href='ServiceDeskTeam.aspx']")))).click().perform()
            if req.get()==1:
                WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//option[@value='Password Reset']"))).click()
            elif req.get()==2:
                WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//option[@selected='selected'][contains(text(),'Access Request')]"))).click()
            ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@id='ctl00_ContentPlaceHolder1_txtSearch']")))).click().send_keys(accessNo.get()).send_keys(Keys.RETURN).perform()
            try:
                cursor = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//span[@id='ctl00_ContentPlaceHolder1_lblstatus'] | //span[@id='ctl00_ContentPlaceHolder1_lblpwdupdate']")))

                if cursor.text == "Ticket No.: "+accessNo.get()+" not found":
                    msgf6.set("You have provided incorrect request number. Can you please verify all the details.")
                    driver.close()
                    driver.switch_to_window(window_before)
                    accessNo.set("")
                    req.set(0)
                    f6.tkraise()
                elif cursor.text=="":
                    raise Exception
            except:
                cursor = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//*[@id='ctl00_ContentPlaceHolder1_dgvdetail']/tbody/tr[2]/td[10]| //*[@id='ctl00_ContentPlaceHolder1_gvpwdreset']/tbody/tr[2]/td[9]")))
                if "Approved" in cursor.text.strip() or "Submitted" in cursor.text.strip():
                    ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@id='ctl00_ContentPlaceHolder1_dgvdetail_ctl02_chkupdate'] | //input[@id='ctl00_ContentPlaceHolder1_gvpwdreset_ctl02_chkupdatepwd']")))).click().perform()
                    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@id='ctl00_ContentPlaceHolder1_dgvdetail_ctl02_fupload']|//input[@id='ctl00_ContentPlaceHolder1_gvpwdreset_ctl02_FileUploadpwd']"))).send_keys(path)
                    sleep(2)
                    ActionChains(driver).move_to_element(WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//input[@id='ctl00_ContentPlaceHolder1_btnupdate']|//input[@id='ctl00_ContentPlaceHolder1_btnupdatepwd']")))).click().perform()
                    driver.close()
                    driver.switch_to_window(window_before)
                    initial_screen()
                else:
                    msgf6.set("Request is not in approved state. Please get it approved and upload manually. Please press logout .")
                    driver.close()
                    driver.switch_to_window(window_before)
                    accessNo.set("")
                    req.set(0)
                    port.set(2)
                    portal()
                    f6.tkraise()

    elif port.get() == 2:
        initial_screen()

def logout():
    global driver
    driver.switch_to.default_content()
    cursor = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//frame[@title='Banner frame']")))
    driver.switch_to.frame(cursor)
    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//a[@title='Logout' and contains(text(),'Logout')]"))).click()

            
def initial_screen():
    logout()
    url.set("1")
    msgf1.set("Note:If both username and password field is left empty then default ID(L096535) will be used to login")
    msgf2.set("Please enter Lid or LId's separated by either space,comma(,) or enter.")
    msgf3.set("")
    msgf5.set("")
    msgf6.set("")
    msgf4.set("")
    pass_change.set("")
    val.set(0)
    user_lid.set("")
    user_password.set("")
    model_id.set("")
    app_name.set("")
    req.set(0)
    port.set(0)
    accessNo.set("")
    userEnteredRoles.delete('1.0', 'end-1c')
    mod.set(0)
    rem.set(0)
    app_name_remove.set("")
    rolesRemoved.delete('1.0', 'end-1c')
    lid.delete('1.0', 'end-1c')
    lbl1.grid_forget()
    lbl2.grid_forget()
    app_label_remove.grid_forget()
    app_entry_remove.grid_forget()
    ent1.grid_forget()
    ent2.grid_forget()
    lbl3.grid_forget()
    roles_label.grid_forget()
    rolesRemoved.grid_forget()
    access_no_lbl.grid_forget()
    access_no_ent.grid_forget()
    pass_req.grid_forget()
    acc_req.grid_forget()
    btn.grid_forget()
    userEnteredRoles.grid_forget()
    f1.tkraise()
    
def is_password():
    if val.get()==2:
        initial_screen()
    elif val.get()==3:
        f5.tkraise()
        


def cont():
    global lid
    if lid.get("1.0",'end-1c')=="":
        msgf2.set("You did not provide required information. Please provide LID's in the below box")
    else:
        if val.get()==1:
            f5.tkraise()
        elif val.get()==2 or val.get()==3:
            f3.tkraise()        
        elif val.get()==4:
            f4.tkraise()
        else:
            msgf2.set("Please select on of the four options.Please enter Lid or LId's separated by either space,comma(,) or enter.")
    

def model():
    if mod.get() == 1:
        msgf3.set("")
        lbl3.grid_forget()
        userEnteredRoles.grid_forget()
        lbl1.grid(row=3,column=0,sticky=tk.W)
        ent1.grid(row=3,column=1,sticky=tk.E)
        lbl2.grid(row=4,column=0,sticky=tk.W)
        ent2.grid(row=4,column=1,sticky=tk.E)
    elif mod.get() == 2:
        lbl1.grid_forget()
        lbl2.grid_forget()
        ent1.grid_forget()
        ent2.grid_forget()
        msgf3.set("")
        lbl3.grid(row=3,columnspan=2,padx = 10, sticky="NSEW")
        userEnteredRoles.grid(row=4,columnspan=2,sticky="NSEW",padx=(5,10))
    
def portal():
    btn.grid(row=5,column=0,columnspan=2,ipadx=30,padx=20,pady=10,sticky=tk.W)
    if port.get() == 1:
        msgf6.set("What type of request is it?")
        access_no_lbl.grid(row=2,column=0,padx=20,pady=10,sticky=tk.W)
        access_no_ent.grid(row=2,column=1,padx=20,pady=10,sticky=tk.E)
        pass_req.grid(row=3,column=0,padx=20,pady=10,sticky=tk.W)
        acc_req.grid(row=3,column=1,padx=20,pady=10,sticky=tk.E)
        btn_value.set("Continue")
        
    elif port.get() == 2:
        msgf6.set("")
        access_no_lbl.grid_forget()
        access_no_ent.grid_forget()
        pass_req.grid_forget()
        acc_req.grid_forget()        
        btn_value.set("Logout")
        

def select_for_removal():
    if rem.get() == 1:
        msgf4.set("")
        roles_label.grid_forget()
        rolesRemoved.grid_forget()
        app_label_remove.grid(row=3,columnspan=2,padx = 10, sticky="NSEW")
        app_entry_remove.grid(row=4,columnspan=2,sticky="NSEW",padx=10)
        
    elif rem.get() == 2:
        msgf4.set("")
        roles_label.grid(row=3,columnspan=2,padx = 10, sticky="NSEW")
        rolesRemoved.grid(row=4,columnspan=2,sticky="NSEW",padx=(5,10))
        app_label_remove.grid_forget()
        app_entry_remove.grid_forget()
    


    
root=tk.Tk()
root.geometry('360x230')
root.resizable(0,0)
root.iconbitmap(r'P:\imran-TEMS\password\icon.ico')
#variables
url = StringVar(value="1")
msgf1= StringVar()
msgf2 =StringVar()
msgf3 =StringVar()
msgf4 = StringVar()
msgf5 =StringVar()
msgf6 =StringVar()
pass_change =StringVar()

val=IntVar(value=0)
port = IntVar(value=0)
req = IntVar(value=0)
rem = IntVar(value=0)
root.title("Password Reset Tool")
user_lid = StringVar()
lid =StringVar()
user_password = StringVar()
model_id= StringVar()
app_name= StringVar()
app_name_remove= StringVar()
userEnteredRoles= StringVar()
rolesRemoved = StringVar()
accessNo = StringVar()
mod = IntVar(value=0)
btn_value = StringVar(value="Continue")


chrome_options=Options()
#chrome_options.add_argument("--headless")
chrome_options.add_argument("start-maximized")
try:
    driver = webdriver.Chrome(chrome_options=chrome_options,executable_path=r"P:\imran-TEMS\selenium-3.4.3\chromedriver.exe")
except:
    driver = webdriver.Chrome(chrome_options=chrome_options,executable_path=r"P:\imran-TEMS\selenium-3.4.3\chromedriver-2.33.exe")
f1 = tk.Frame(root)
f2 = tk.Frame(root)
f3 = tk.Frame(root)
f4 = tk.Frame(root)
f5 = tk.Frame(root)
f6 = tk.Frame(root)
#part of frame 3
lbl1=tk.Label(f3,text="Enter Model Id")
ent1=tk.Entry(f3,textvariable = model_id)
lbl2=tk.Label(f3,text="Enter Application name(ex: sol,ce or all)")
ent2=tk.Entry(f3,textvariable=app_name)
lbl3 = tk.Label(f3,text="Please enter exact role name('s) separated by either space,comma(,) or enter." ,wraplength=340,justify=tk.LEFT)
userEnteredRoles = tkscrolled.ScrolledText(f3,width=45,height=4)

#part of Frame 6
access_no_lbl = tk.Label(f6,text="Please enter access number")
access_no_ent = tk.Entry(f6,textvariable = accessNo)
pass_req= tk.Radiobutton(f6,text="Password Reset",variable=req,value=1)
acc_req = tk.Radiobutton(f6,text="Access Request",variable=req,value=2)
btn=tk.Button(f6, textvariable=btn_value,command=access_request_portal)

#part of Frame 4
roles_label = tk.Label(f4,text="Please enter exact role name('s) separated by either space,comma(,) or enter." ,wraplength=340,justify=tk.LEFT)
rolesRemoved = tkscrolled.ScrolledText(f4,width=45,height=4)
app_label_remove = tk.Label(f4,text="Enter Application name(ex: sol,ce,cis):",wraplength=340,justify=tk.LEFT)
app_entry_remove = tk.Entry(f4,textvariable = app_name_remove)


for frame in (f1, f2, f3, f4,f5,f6):
    frame.grid(row=0, column=0, sticky='news')




create_frames()
f1.tkraise()
root.mainloop()


