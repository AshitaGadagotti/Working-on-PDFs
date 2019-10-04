##Step1: Import required modules

from selenium import webdriver #Allows you to launch/initialize a browser
from selenium.webdriver.common.by import By #Allows you to search for things using specific parameters
from selenium.webdriver.support.ui import WebDriverWait #Allows you to wait for a page to load
from selenium.webdriver.support import expected_conditions as EC #Specify what you are looking for on a specific page in order to determine that the webpage has loaded.
from selenium.common.exceptions import TimeoutException #Handling a timeout situation.
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import os
from PyPDF2 import PdfFileReader, PdfFileWriter
import re
import fitz

#KYC_Ref_Number list

import pandas as pd
file_location = r"C:\Users\user\Desktop\xyz.xlsm"
KYC_list = pd.read_excel(file_location)
List = KYC_list['KYC_ref_number'].tolist()
List = list(map(str,List))
print (List)

#Create new browsing session and specify the download directory

chrome_options = Options()
chrome_options.add_argument("start-maximized")
prefs = {"profile.default_content_settings.popups": 0,
                 "download.default_directory": r"C:\Users\user\MI\Reports\\", # IMPORTANT - ENDING SLASH IS IMPORTANT
                 "directory_upgrade": True}
chrome_options.add_experimental_option("prefs", prefs)
browser = webdriver.Chrome("C:/Users/ashita.gadagotti/Downloads/chromedriver_win32/chromedriver.exe",chrome_options=chrome_options) 

#Launch URL

browser.get("https://xyz.com/")

#Login Credentials

username_input = browser.find_element_by_xpath("//input[@id='Email']").send_keys("abc@gmail.com")
password_input = browser.find_element_by_xpath("//input[@id='Password']").send_keys("@@@@@")
browser.find_element_by_xpath("//button[@class='btn btn-lg btn-primary btn-block']").click()
browser.implicitly_wait(10)

#Customersearch Page & iterate through each Reference ID

browser.find_element_by_xpath("//a[contains(text(),'Customer search')]").click()
for id in List:
          ID_entry = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@id='Name']")))
          ID_entry.send_keys(id)
          browser.find_element_by_xpath("//select[@id='SearchOn']//option[contains(text(),'Reference')]").click()
          browser.find_element_by_xpath("//button[@class='btn btn-primary']").click()
          browser.implicitly_wait(5)
          browser.find_element_by_xpath("/html[1]/body[1]/div[2]/div[1]/div[1]/div[4]/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[1]/a[1]").click()
          browser.implicitly_wait(10)
          browser.find_element_by_xpath("//button[@class='btn btn-primary']").click()
          browser.back()
          WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@id='Name']")))
          browser.find_element_by_xpath("//input[@id='Name']").clear();
          browser.implicitly_wait(10)


#Step2: Renaming Files
#open the individual pdf of downloaded files

def rename_pdfs():
  dirList=os.listdir('C:/Users/user/MI/Reports')
  for fname in dirList:
     if fname.endswith(".pdf"):
         # open the individual pdf
         pdf = "C:/Users/user/MI/Reports/" + str(fname)
         pdf_reader = fitz.open(pdf)

         # access the individual page
         page_obj = pdf_reader.loadPage(0)
         # extract the the text
         pdf_text = page_obj.getText()
         
         # use regex to find information
         r1 = re.search(r"(?<!\d)\d{6}(?!\d)",pdf_text)
         ref = r1.group(0)

         pdf_reader.close()

         # rename the pdf based on the information in the pdf
         os.rename(pdf, "C:/Users/user/MI/Reports/"  + ref + ".pdf")

rename_pdfs()

#Step3:Splitting Monthly BatchPages

def pdf_splitter(fname):
  mypath = 'C:/Users/user/Downloads'
  fname = mypath + "/" + "Batch September 2019.pdf"
  despath = 'C:/Users/user/MI/Reports/Batch pdfs' + "/"
  
  pdf = PdfFileReader(fname)
  for page in range(pdf.getNumPages()):
    pdf_writer = PdfFileWriter()
    pdf_writer.addPage(pdf.getPage(page))

    output_filename = '{}_page_{}.pdf'.format(
        despath, page+1)

    with open(output_filename, 'wb') as out:
        pdf_writer.write(out)

    print('Created: {}'.format(output_filename))

  import glob

  if __name__ == '__main__':
    paths = glob.glob('*.pdf')
    for fname in paths:
        pdf_splitter(fname)

pdf_splitter("Batch September 2019.pdf")

#Step4:open the individual pdf of batchpages and rename

import os
from PyPDF2 import PdfFileReader, PdfFileWriter
import re
import fitz
def rename_pdfs1():
  dirList=os.listdir('C:/Users/user/MI/Reports/Batch pdfs')
  for fname in dirList:
     if fname.endswith(".pdf"):
         # open the individual pdf
         pdf = "C:/Users/user/MI/Reports/Batch pdfs" + str(fname)
         pdf_reader = fitz.open(pdf)

         # access the individual page
         page_obj = pdf_reader.loadPage(0)
         # extract the the text
         pdf_text = page_obj.getText()
         
         # use regex to find information  
         r1 = re.search('number: (.*) ',pdf_text)
         ref = r1.group(1)
            
         pdf_reader.close()

         # rename the pdf based on the information in the pdf
         os.rename(pdf, "C:/Users/user/MI/Reports/Batch pdfs"  + ref + ".pdf")

rename_pdfs1()

#Step5:Merging Files

import os
from PyPDF2 import PdfFileReader, PdfFileWriter
import re
import fitz

dirList=os.listdir('C:/Users/user/MI/Reports')
Finalpages = fitz.open('C:/Users/auser/Downloads/Pages 3 & 4.pdf')

for fname in dirList:
    if fname.endswith(".pdf"):
        print ("Processing " + str(fname) + "....")

 # set up the pdfwriter and the input source
        output = fitz.open()
        inputFile = "C:/Users/user/MI/Reports/Batch pdfs/" + str(fname)
        
        input1 = fitz.open(inputFile)

        # add the candidate report page to the PDF output
        output.insertPDF(input1)

        # add the riskscreen page to the PDF output
        riskscreenpage = "C:/Users/user/MI/Reports/" + str(fname)
        input2 = fitz.open(riskscreenpage)

        output.insertPDF(input2,to_page=0)
       
        # add the rest of the pages to the PDf output
        output.insertPDF(Finalpages)

        pdf_reader = fitz.open(inputFile)
        page_obj = pdf_reader.loadPage(0)
        pdf_text = page_obj.getText()
         
         # use regex to find information
        r = re.search('name: (.*) ',pdf_text)
        candidate = r.group(1)

        # write the output file - change output folder as needed
        outputFile = "C:/Users/user/MI/Reports/" + candidate + "_" + str(fname)
        output.save(outputFile)


        print ("Conversion complete.")


