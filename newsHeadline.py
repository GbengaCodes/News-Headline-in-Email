from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import datetime, time, os
import smtplib
from email.message import EmailMessage
import os
import win32com.client as win32


print("- ->  Please, open Outlook App now!")
#getting current date
td = datetime.date.today()
dt=td.strftime("%b-%d-%Y")


# create a webdriver object for chrome-option and configure
wait_imp = 10
CO = webdriver.ChromeOptions()
CO.add_experimental_option('useAutomationExtension', False)
CO.add_argument('--ignore-certificate-errors')
CO.add_argument('--start-maximized')
wd = webdriver.Chrome(r'C:\Users\ojotg\OneDrive\Documents\Career\learningResources\chromedriver.exe',options=CO)


#updating user on background process
print ("- ->    Connecting to Authentic News source, Please wait .....\n")
news_source = "https://news.google.com/topstories?hl=en-US&gl=US&ceid=US:en"


#displaying news information to user on console
print (" ------------------------------------------------------------------------------------------- ")
print (">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  TODAY's TOP NEWS HEADLINES  <<<<<<<<<<<<<<<<<<<<<<<<<<<<< ")
print ("Date:",td.strftime("%b-%d-%Y"))
print ("-------------------------- ")


#getting necessary web elements from the new source   
wd.get(news_source)
wd.implicitly_wait(wait_imp)
elems = wd.find_elements(By.TAG_NAME, 'h3')


#writing news info into a file called "newsinfo" in the same directory as source file
file_loc = r'C:\Users\ojotg\OneDrive\Documents\Career\Projects\newsHeadline\newsinfo.txt'
file_to_write = open(file_loc, 'w+')
file_to_write.write("TODAY's TOP NEWS HEADLINES \n")
file_to_write.write("Date: "+ dt+"\n")
ind = 1
for elem in elems:
    file_to_write.write(str(ind)+ '>> ')
    file_to_write.write(elem.text+'\n')
    print (str(ind) + ") " + elem.text)
    ind += 1
file_to_write.close()
print('\n')


#sending email to recipient with "newsinfo.txt" attached to email
olApp=win32.Dispatch("Outlook.Application")
olNS=olApp.GetNameSpace('MAPI')
mailItem=olApp.CreateItem(0)
mailItem.Subject='News Headline'
mailItem.BodyFormat=1
mailItem.Body="Find the attached document for detailed NEWS .. "
mailItem.To="<Recipient E-mail"
mailItem.Attachments.Add(os.path.join(os.getcwd(), 'newsinfo.txt'))
mailItem.Display()
time.sleep(2)
mailItem.Send()
print("- ->     Check inbox for Today's News")
