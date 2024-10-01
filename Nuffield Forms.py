# Creates Multiple word documents for Nuffield using a word template
# v1.1 23/02/23
# Change to match new export activities button code
# v1.2 15/06/23  - Change to writeupp report
# v1.3 16/08/23 Change in writeupp
# v1.5 01/05/24 Change in Writeupp and Selenium 4 support
# v1.6 01/10/24 Change to CSV format


import os
import csv
import time
import glob
from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from tkinter import *
import pymsgbox
from docxtpl import DocxTemplate
from selenium.webdriver.firefox.service import Service

def loginWriteupp():
    #testfield = entrySuffix.get()f
    #Login to writeUpp
    loginDriver = driver
    loginDriver.get(loginURL)
    time.sleep(2)
    userNameField = loginDriver.find_element(By.ID,'EmailInput')
    userNameField.send_keys(userName)
    passwordField = driver.find_element(By.ID,'PasswordInput')
    passwordField.send_keys(password)
    time.sleep(1)
    submitButton = driver.find_element(By.XPATH,'/html/body/section/div[2]/form/button')
    submitButton.click()
    time.sleep(5)

def getInsuranceCompanies():
    #Find all the third party insurance companies and export them and then read into a list
    driver.get(thirdURL)
    time.sleep(2)
    insurerSelect = driver.find_element(By.XPATH,'/html/body/form/div[5]/div/div/div[2]/select')
    Select(insurerSelect).select_by_visible_text('Insurer')
    time.sleep(2)
    searchButton = driver.find_element(By.XPATH,'/html/body/form/div[5]/div/div/div[3]/button')
    searchButton.click()
    time.sleep(2)
    exportCSV = driver.find_element(By.XPATH, "/html/body/form/div[5]/div/div/div[4]/div/div/div/div/div/div[3]/button")
    exportCSV.click()
    time.sleep(2)
    os.chdir(wd)
    global companies
    companies = [] #this list contains all the insurance companies in use.
    thirdPartiesFile = open('ThirdParties.csv')
    csv_insurers = csv.reader(thirdPartiesFile)
    for row in csv_insurers:
        companies.append(row[0])
    companies.pop(0) # get rid of header row
    thirdPartiesFile.close()

def process_patients():
    os.chdir(wd)
    with open(activity_filename, 'r', encoding='utf-8') as p:
        patients = csv.reader(p)
        next(patients) # skip header row
        for patient in patients:
            tp_name = patient[1]
            tp_appointment_time = patient[5]
            tp_appointment_date = patient[7]
            tp_appointment_type = patient[2]
            wuid = patient[0]
            searchField = driver.find_element(By.ID,'ctl00_ctl00_Content_siteHead_dfSearchWidget')
            searchField.send_keys(wuid)
            driver.find_element(By.ID,'ctl00_ctl00_Content_siteHead_btnSearch').click()
            time.sleep(1)
            age_dob_field = driver.find_element(By.ID,'ctl00_ctl00_Content_ContentPlaceHolderPS_dateOfBirth')
            age_dob = age_dob_field.text
            tp_DOB = age_dob[:10]
            tp_age = age_dob[age_dob.find("(")+1:age_dob.find(")")] #grab text between the brackets
            try:
                tp_address = driver.find_element(By.XPATH,'/html/body/form/div[4]/div[3]/div/div/div[1]/article[1]/table/tbody/tr[3]/td').text
                tp_home_phone = driver.find_element(By.XPATH,
                    '/html/body/form/div[4]/div[3]/div/div/div[2]/article[1]/table/tbody/tr[1]/td/span/a').text
                tp_mobile = driver.find_element(By.XPATH,
                    '/html/body/form/div[4]/div[3]/div/div/div[2]/article[1]/table/tbody/tr[3]/td/span/a').text
                tp_email = driver.find_element_by_xpath(
                    '/html/body/form/div[4]/div[3]/div/div/div[2]/article[1]/table/tbody/tr[4]/td/span/a').text
                tp_nhs = driver.find_element(By.XPATH,
                    '/html/body/form/div[4]/div[3]/div/div/div[1]/article[1]/table/tbody/tr[8]/td/div/p').text
            except:
                tp_address = driver.find_element(By.XPATH,
                    '/html/body/form/div[5]/div[3]/div/div/div[1]/article[1]/table/tbody/tr[3]/td').text

                tp_home_phone = driver.find_element(By.XPATH,'/html/body/form/div[5]/div[3]/div/div/div[2]/article[1]/table/tbody/tr[1]/td/span/a').text
                tp_mobile = driver.find_element(By.XPATH,'/html/body/form/div[5]/div[3]/div/div/div[2]/article[1]/table/tbody/tr[3]/td/span/a').text
                tp_email = driver.find_element(By.XPATH,'/html/body/form/div[5]/div[3]/div/div/div[2]/article[1]/table/tbody/tr[4]/td/span/a').text
                tp_nhs = driver.find_element(By.XPATH,'/html/body/form/div[5]/div[3]/div/div/div[1]/article[1]/table/tbody/tr[8]/td/div/p').text

            #Get GP and insurance details
            third_parties = driver.find_elements(By.CLASS_NAME,'patient-summary__third-parties__name')
            thirdparty_attributes = driver.find_elements(By.CLASS_NAME,'patient-summary__third-parties__attribute')
            #set theses tom blank in case they don't exist
            tp_insurance_co = ''
            tp_insurance_co_address =''
            tp_membership_no =''
            # Remove extra text from the address
            tp_address = tp_address.replace('Show on Google Map', '')


            tp_gp_surgery =''
            tp_authorisation =''
            for third_party in third_parties:
                third_party_text = third_party.text
                first_word = third_party_text.split(' ', 1)[0]
                back_string = third_party_text.split("- ", 1)[1]
                doctor_word = back_string.split(' ', 1)[0]
                first_word_lower = first_word.lower()
                for company in companies:
                    company_lower = company.lower()
                    if first_word_lower in company_lower:
                        tp_insurance_co = third_party_text.split('-', 1)[0]
                        tp_insurance_co_address = back_string
                    elif doctor_word in dr_list:
                    # print('found a Doctor')
                    # print(third_party_text)
                        tp_gp_name = back_string.split(',', 1)[0]
                        tp_gp_surgery = third_party_text.replace(tp_gp_name+",","") #remove doctors name
                        tp_surgery_name = tp_gp_surgery.split('-',1)[0]
                    else:
                         print('You have found something else')
            #Check for policy number and autorisation code
                    if thirdparty_attributes != []:
                        # print('something here')
                        for attribute in thirdparty_attributes:
                            third_party_attribute_text = attribute.text
                            if "Policy Number" in third_party_attribute_text:
                                tp_membership_no = third_party_attribute_text.split(':', 1)[1] #get the text after the :
                            if "Authorisation Code" in third_party_attribute_text:
                                tp_authorisation = third_party_attribute_text.split(':', 1)[1] #get the text after the :

            # Start creating the spreadsheet for this patient
            os.chdir(wd)
            patient_file = DocxTemplate(template_file)
            # Populate the dictionary
            context = {
                'f_NHS_no': tp_nhs,
                'f_title': tp_name.split(' ', 1)[0],
                'f_name': tp_name.split(' ', 1)[1],
                'f_DOB': tp_DOB,
                'f_mobile': tp_mobile,
                'f_landline': tp_home_phone,
                'f_address': tp_address,
                'f_email': tp_email,
                'f_GP_details': tp_gp_name + ' ' + tp_surgery_name,
                'f_insurance': tp_insurance_co,
                'f_membership': tp_membership_no,
                'f_authorisation': tp_authorisation,
                'f_appointment_date': tp_appointment_date,
                'f_appointment_time': tp_appointment_time
            }

            patient_file.render(context)
            this_filename = tp_name+'.docx'
            os.chdir(this_dir)
            patient_file.save(this_filename)




def finishUp():
    root.destroy()
    # Delete all oldfiles
    oldFiles = glob.glob(wd + '//*.csv')
    for f in oldFiles:
        os.remove(f)
    return

def alldone():
    pymsgbox.alert('All done')
    driver.quit()
    return

def getActivity():
    #Export all the Activity for the given dates in WriteUpp

    #Delete all oldfiles
    oldFiles = glob.glob(wd+'//*.csv')
    for f in oldFiles:
        os.remove(f)

    # Get the Acivity report
    root.withdraw()
    driver.get(activityURL)
    time.sleep(1)
    pymsgbox.alert('Enter Dates and  click OK')

    time.sleep(1)
    export_button = driver.find_element(By.XPATH, "/html/body/form/div[5]/div/div/div[6]/div/div/div/div/div/div/div")
    export_button.click()
    time.sleep(1)
    export_csv_button = driver.find_element(By.XPATH, "/html/body/form/div[5]/div/div/div[6]/div/div/div/div/div/div/div/ul/li[2]/button")
    export_csv_button.click()
    time.sleep(3) # increased as ile sometimes not there at point of rename
    os.chdir(wd)
    os.rename(wu_activity_filename,activity_filename)

def setup_folder():
    folder_name = entryFolder.get()
    global this_dir
    this_dir = os.path.join(HospitalSheetDirectory, folder_name)
    if not os.path.exists(this_dir):
        os.mkdir(this_dir)
    return

version_no = "v1.6 JB 08/07/24"
writeUppURL = 'https://dr-emma-howard-dermatology.writeupp.com/'
driverPath = "C:\\Users\\justi\\Dropbox\\PC (2)\\Desktop\\Clinics\\geckodriver.exe"
thirdURL = writeUppURL + '/admin/thirdparties.aspx'
loginURL = 'https://portal.writeupp.com/login'
patientsURL = writeUppURL + '/admin/data-management/patients.aspx'
activityURL = writeUppURL + '/contactsbydate.aspx'
patientsByInsurer = writeUppURL + '/patientsbythirdparty.aspx'
userName = 'aliwid5@gmail.com'
password = 'Melanoma1!'
testWUID = 'WU1191771'
wd = 'C:\\Users\\justi\\Dropbox\\PC (2)\\Desktop\\Clinics'
HospitalSheetDirectory = wd+'\\Nuffield Cheltenham'
template_file = 'Patient Booking Form.docx'
downloadDirectory = wd
dr_list = ['Doctor','Dr','Dr.','Doctor,','Dr,','Dr.,','General']
wu_activity_filename = 'Activity by date.csv'
activity_filename = 'Activity.csv'



firefox_location = r"C:\Program Files\Mozilla Firefox\firefox.exe"

firefox_options = Options()
firefox_options.binary_location = firefox_location
firefox_options.add_argument("--disable-infobars")
firefox_options.add_argument("--disable-extensions")
firefox_options.add_argument("--disable-popup-blocking")
# Set the download directory preference (Assuming you have defined downloadDirectory somewhere in your code)
firefox_options.set_preference('browser.download.dir', downloadDirectory)
firefox_options.set_preference('browser.download.folderList', 2)
firefox_options.set_preference('browser.download.manager.showWhenStarting', False)
firefox_options.set_preference('browser.helperApps.neverAsk.saveToDisk', 'text/csv')
firefox_options.set_preference('browser.download.panel.shown', False)
firefox_options.set_preference('browser.download.manager.showAlertOnComplete', False)

# Update the driver path (Assuming you have defined driverPath somewhere in your code)
driver_path = driverPath

driver = webdriver.Firefox(
    service=Service(driver_path),  # Pass the service object with the driver path
    options=firefox_options
)

driver.implicitly_wait(5)

def goforit():
    loginWriteupp()
    setup_folder()
    getActivity()
    getInsuranceCompanies()
    process_patients()
    finishUp()
    alldone()


# GUI It all starts here
# Get the login password and billing file suffix
root = Tk()

label_1 = Label(root, text = 'Password:')
label_4 = Label(root, text = 'Clinic Folder' )
label_5 = Label(root, text = version_no)

entryPassword = Entry(root)
entryFolder = Entry(root)

label_1.grid(row=0)
label_4.grid(row=3)
label_5.grid(row=5)

entryPassword.grid(row=0,column = 1)
entryPassword.insert(0,password) # Display the password in the code
entryFolder.grid(row=3,column = 1)

submitButton1 = Button(root, text = 'Press to submit',command=goforit) #live
#submitButton1 = Button(root, text = 'Press to submit',command=testrun) #test

submitButton1.grid(row=4,column=1)

root.mainloop()