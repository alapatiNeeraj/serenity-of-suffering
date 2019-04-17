import win32com.client
import xlrd
from datetime import date,timedelta
from pywintypes import com_error
from bs4 import BeautifulSoup as bs
import re
from win32com.mapi.mapitags import PROP_TAG, PT_UNICODE

""" Everyday in the evening TEMS offshore creates all the reports that needs to be sent by onshore counter on the next day adn saved in draft folder. Reports are prepared in such a way
that everything is update with date and to and BCC addresses."""


def suffix(d):
    return 'th' if 11<=d<=13 else {1:'st',2:'nd',3:'rd'}.get(d%10,'th')

#This below function will take the mail and make necessary changes and save it in the draft folder
def save_draft(msg,to_bcc):
    #Reading body of the mail in HTML Format
    body_content = msg.HTMLBody

    #If the mail contains images then that image gets broken while saving in draft. So saved images at below location so it will attach from here
    path_of_img = {0:'P:\imran-TEMS\dummy\Tems.png',1:'P:\imran-TEMS\dummy\westpaclogo.jpg'}
    count=0
    num = 0

    #Below is the partial pattern that we will be searching in the mail to replace broken image using regular expression
    ext = {0:".png@.{8}..{8}",1:".jpg@.{8}..{8}"}

    #Below loop is to find the pattern and replace the pattern in the mail body
    while count<2 and num<50:
        #Below is the complete pattern of the broken image. For ex: cid:image001.png@8CE5GIN7.85GHY26I
        pat = "cid:image00"+str(num)+ext[count]

        match = re.search(pat,body_content)

        #Replacing the pattern to image path 
        body_content = re.sub(pat,path_of_img[count],body_content)
        num+=1
        if match is None:
            continue
        count+=1
    body_content = re.sub("westpac-logo",path_of_img[1],body_content)

    #using beautifulSoup module to read contents of HTML body of email
    rsoup  = bs(body_content, "html.parser")
    nodes  = rsoup.find('div',{'class':'WordSection1'})
    #Reading all the elements of table 
    nodes1 = nodes.find('table').find_all('tr')
    
    #Below block is to get todays date and future date
    today_date = date.today()
    tomorrow_date = date.today()+timedelta(days=3) if today_date.weekday() == 4 else date.today()+timedelta(days=1)
    
    todays_date = today_date.strftime('{S} %B %Y, %A').replace('{S}',str(today_date.day)+suffix(today_date.day))
    todays_date_no_year = today_date.strftime('{S} %B, %A').replace('{S}',str(today_date.day)+suffix(today_date.day))
    
    tomorrows_date = tomorrow_date.strftime('{S} %B %Y, %A').replace('{S}',str(tomorrow_date.day)+suffix(tomorrow_date.day))
    
    try:
        #Below for loop is the iterate through all the contents of table and find date column and replace the date
        for node in nodes1:     
            for d in node.findAll('td'):
                email_date  = list(filter(None, d.get_text().replace('\xa0',"").strip().split(" ")))
                if email_date == todays_date_no_year.split(" ") or email_date == todays_date.split(" "):
                    for data in d:                       
                        p_tag = rsoup.new_tag("p")
                        p_tag["align"] = "center"
                        p_tag["class"] = "MsoNormal"
                        p_tag["style"] ="text-align:center;mso-element:frame;mso-element-frame-hspace:38.7pt;mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal:column;mso-height-rule:exactly"

                        span_tag = rsoup.new_tag("span")
                        span_tag["style"] = "font-size:9.0pt;mso-fareast-language:EN-AU"
                        span_tag.string = tomorrows_date
                        
                        b_tag = rsoup.new_tag("b")
                        new_tag = rsoup.new_tag("o:p")

                        b_tag.append(span_tag)
                        p_tag.append(b_tag)
                        data.replace_with(p_tag)
                    raise StopIteration
    except StopIteration:
        pass

    #Below block is to create outllok item, then draft a mail and move it to draft folder
    mail = outlook.CreateItem(0)
    mail.To = 'TEMS Service Desk'
    mail.bcc = to_bcc
    subject  = msg.subject.split(" ")
    subject[-1] = str(tomorrow_date.year)
    subject[-2] = tomorrow_date.strftime('%B')
    subject[-3] = str(tomorrow_date.day)+suffix(tomorrow_date.day)
    mail.Subject = " ".join(subject)
    mail.HTMLBody = rsoup.prettify()
    mail.Move(draft_folder)



# Below we are reading the mails and accessing draft foler and health check folder
# Read the health check mails of present day that are present in Health Check folder and then save it to draft foler after making necesary changes
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
draft_folder = outlook.GetNamespace("MAPI").Folders['TEMS Service Desk'].Folders['Drafts']
health_check_folder = namespace.Folders['TEMS Service Desk'].Folders['Inbox'].Folders['Health Check Report']

#Excel file is created with applcation name and BCC field in each column.
#Note: This program takes the mail items as per excel so if any new health check mail that needs to be added, user can just add the details of that health check in the excel sheet. Same can be done to remove mail just delete the mail details from excel sheet
xls = xlrd.open_workbook(r'P:\imran-TEMS\dummy\BCC.xlsx')
sheet = xls.sheet_by_index(0)
#application list 
app_list = [sheet.cell_value(row, 0).upper() for row in range(sheet.nrows)]
#BCC details
bcc_list = [sheet.cell_value(row, 1) for row in range(sheet.nrows)]
BCC = dict(zip(app_list,bcc_list))

#Reading all the messages in health check folder
messages = health_check_folder.Items

#creating copy of application list that we read from Excel
check_list = app_list[:]

#Iterating each mail present in the health check folder
for message in messages:
    #Below is the condition for the program to exit. If check_list variable that was created above has no elements or if date is not today's date then the progam will stop
    if len(check_list) == 0 or  message.ReceivedTime.date() !=  date.today():
        break

    #Splitting subject of the current mail and taking only first word of the  subject
    app_name = message.subject.upper().split(" ")[0]

    #if the app_name is present in check_list variable then perform below 2 steps:
    #      1) Remove the app_name from check_list variable
    #      2) execute save_draft function
    if app_name in check_list:
        check_list.remove(app_name)
        save_draft(message,BCC[app_name]) #passing mail contents, BCC 


