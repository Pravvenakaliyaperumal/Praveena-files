from datetime import time
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import Service

import os
import xlrd

from selenium.webdriver.common.by import By

location="C:\\Users\\CoxPHIT\\OneDrive\\Documents\\try_excel_xlss.xls"
var=xlrd.open_workbook(location)
sht= var.sheet_by_index(0)
submitted_client=[]
not_submitted_client=[]
start_count=139
count=150
for row_count in range(start_count,count+1):
    print(row_count)
    madicaid_num=sht.cell_value(row_count,10)
    Type_of_service=sht.cell_value(row_count,16)
    option=Options()
    option.add_argument("--window-size=1920,1080")
    driver= webdriver.Chrome(options=option)
    
    driver.get("https://www.njmmis.com/default.aspx")
    xpath="//*[@class='nav-link siteNavItem'] [contains(text(),'Login')][2]"
    
    driver.find_element(By.XPATH, xpath).click()
    time.sleep(5)
    
    xpath_username='//*[@id="txtUserName"]'
    
    driver.find_element(By.XPATH,xpath_username).click()
    driver.find_element(By.XPATH,xpath_username).click()
    if Type_of_service=="IIC-L" or Type_of_service=="IIC-M" or Type_of_service=="BA":
        loginid="0504271001"
        loginid_pwd = "*******@123"
    elif Type_of_service=="SHR":
        loginid = "0550833002"
        loginid_pwd = "*******@123"
    
    time.sleep(5)
    driver.find_element(By.XPATH,xpath_username).send_keys(loginid)
    time.sleep(3)
    
    xpath_password='//*[@id="txtPassword"]'
    driver.find_element(By.XPATH,xpath_password).click()
    driver.find_element(By.XPATH,xpath_password).click()
    
    time.sleep(5)
    driver.find_element(By.XPATH,xpath_password).send_keys(loginid_pwd)
    time.sleep(3)
    
    xpath_submit='//*[@type="submit"]'
    driver.find_element(By.XPATH,xpath_submit).click()
    time.sleep(5)
    
    xpath_submit_DDE='//*[contains(text(),"Submit DDE Claim")]'
    driver.find_element(By.XPATH,xpath_submit_DDE).click()
    
    time.sleep(5)
    
    
    xpath_checkbox='//*[@id="termschkbx"]'
    authorized=True
    if authorized:
        driver.find_element(By.XPATH, xpath_checkbox).click()
        time.sleep(5)
    
    xpath_CMS="//*[contains(text(),'CMS-1500')]"
    driver.find_element(By.XPATH, xpath_CMS).click()
    time.sleep(2)
    ##################
    
    
    #
    youth_last_name=sht.cell_value(row_count,5)
    youth_first_name=sht.cell_value(row_count,6)
    # Authorization_number=sht.cell_value(row_count,2)
    DOB=sht.cell_value(row_count,8)
    Type_of_service=sht.cell_value(row_count,16)
    dgx=str(sht.cell_value(row_count,7))
    dgx=dgx.replace(".","")
    # dgx=dgx.replace(".","")
    Auth_num=sht.cell_value(row_count,9)
    DOS_start_date=sht.cell_value(row_count,11)
    bill_amount=sht.cell_value(row_count,20)
    Total_units=sht.cell_value(row_count,12)
    qualifier="12"
    current_time = datetime.now()
    formatted_time = current_time.strftime("%m/%d/%y")
    print(type(Type_of_service))
    code,unit_rate=Service.session(Type_of_service)
    code=code.split(".")
    
    total_BLL=float(unit_rate) * float(Total_units)
    print(bill_amount,"bill_amount")
    print(total_BLL,"total_BLL")

    bill_amount=float(bill_amount)
    match=True
    if total_BLL!= bill_amount:
        print("value not matching")
        not_submitted_data=str(youth_last_name)+" "+str(youth_first_name) +" "+str(DOS_start_date)+" Not submitted"
        not_submitted_client.append(not_submitted_data)
        print(not_submitted_client)
        match=True
    else:
        print("its equal")
        match=False
        print(total_BLL)
        NPI=""
        if Type_of_service=="IIC-L" or Type_of_service=="IIC-M" or Type_of_service=="BA":
            NPI="1861865925"
        elif Type_of_service=="SHR":
            NPI="1225547862"
        
        xl_date_DOB=int(DOB)
        
        datetime_date_DOB = xlrd.xldate_as_datetime(xl_date_DOB, 0)
        date_object_DOB = datetime_date_DOB.date()
        print(date_object_DOB.strftime("%m/%d/%Y"))
        DOB_value=date_object_DOB.strftime("%m/%d/%Y")
        print(str(DOB_value))
        print(type(DOB_value))
        
        
        xl_date_DOS_start_date=int(DOS_start_date)
        datetime_date_DOS_start_date = xlrd.xldate_as_datetime(xl_date_DOS_start_date, 0)
        date_object_DOS_start_date = datetime_date_DOS_start_date.date()
        print(date_object_DOS_start_date.strftime("%m/%d/%y"))
        DOS_start_date_value=date_object_DOS_start_date.strftime("%m/%d/%y")
        print(str(DOS_start_date_value))
        
        iframe_id="documentDisplay"
        wait= WebDriverWait(driver, 50)
        
        # Switch to the iframe
        driver.switch_to.frame(iframe_id)
        
        # Now, locate the element inside the iframe
        element = wait.until(EC.element_to_be_clickable((By.ID, "ddeclm1500_txtInsuredIdNumber")))
        
        # After interacting with the element, switch back to the default content
        xpath_insured_id='//*[@id="ddeclm1500_txtInsuredIdNumber"]'
        driver.find_element(By.XPATH, xpath_insured_id).click()
        time.sleep(2)
        driver.find_element(By.XPATH, xpath_insured_id).send_keys(madicaid_num)  #madicaid_num
        time.sleep(2)
        xpath_patient_last_name='//*[@id="ddeclm1500_txtPatientLastName"]'
        driver.find_element(By.XPATH, xpath_patient_last_name).click()
        driver.find_element(By.XPATH, xpath_patient_last_name).send_keys(youth_last_name)   #youth_last_name
        time.sleep(2)
        xpath_patient_first_name='//*[@id="ddeclm1500_txtPatientFirstName"]'
        driver.find_element(By.XPATH, xpath_patient_first_name).click()
        driver.find_element(By.XPATH, xpath_patient_first_name).send_keys(youth_first_name)   #youth_first_name
        time.sleep(2)
        
        xpath_patientDOB='//*[@id="ddeclm1500_txtPatientDOB"]'
        driver.find_element(By.XPATH, xpath_patientDOB).click()
        driver.find_element(By.XPATH, xpath_patientDOB).send_keys((DOB_value))  #DOB need to add
        time.sleep(2)
        
        xpath_diagnx='//*[@id="ddeclm1500_txtRelatedInjury1"]'
        driver.find_element(By.XPATH, xpath_diagnx).click()
        driver.find_element(By.XPATH, xpath_diagnx).send_keys(dgx)   #dgx
        time.sleep(2)
        
        xpath_auth_num='//*[@id="ddeclm1500_txtPriorAuthNumber"]'
        driver.find_element(By.XPATH, xpath_auth_num).click()
        driver.find_element(By.XPATH, xpath_auth_num).send_keys(Auth_num)   #Auth_num
        time.sleep(2)
        
        
        xpath_first_date='//*[@id="ddeclm1500_txtFrom01"]'
        driver.find_element(By.XPATH, xpath_first_date).click()
        driver.find_element(By.XPATH, xpath_first_date).send_keys((DOS_start_date_value))   #DOS
        time.sleep(2)
        
        xpath_qualifier='//*[@id="ddeclm1500_txtPlaceOfService01"]'
        driver.find_element(By.XPATH, xpath_qualifier).click()
        driver.find_element(By.XPATH, xpath_qualifier).send_keys(qualifier)   #qualifier
        time.sleep(2)
        if code[0]:
            xpath_ID_one='//*[@id="ddeclm1500_txtCptHcpcs01"]'
            driver.find_element(By.XPATH, xpath_ID_one).click()
            driver.find_element(By.XPATH, xpath_ID_one).send_keys(code[0])   #qualifier
            time.sleep(2)
        if code[1]:
            xpath_ID_two='//*[@id="ddeclm1500_txtModifier01A"]'
            driver.find_element(By.XPATH, xpath_ID_two).click()
            driver.find_element(By.XPATH, xpath_ID_two).send_keys(code[1])   #qualifier
            time.sleep(2)
        
        if len(code) >2:
            xpath_ID_three='//*[@id="ddeclm1500_txtModifier01B"]'
            driver.find_element(By.XPATH, xpath_ID_three).click()
            driver.find_element(By.XPATH, xpath_ID_three).send_keys(code[2])   #qualifier
            time.sleep(2)
        
        xpath_check_dgnx='//*[@id="ddeclm1500_txtDiag01a"]'
        driver.find_element(By.XPATH, xpath_check_dgnx).click()
        driver.find_element(By.XPATH, xpath_check_dgnx).send_keys("A")   #qualifier
        time.sleep(2)
        
        xpath_charges='//*[@id="ddeclm1500_txtCharges01"]'
        driver.find_element(By.XPATH, xpath_charges).click()
        driver.find_element(By.XPATH, xpath_charges).send_keys("{:.2f}".format(bill_amount))   #bill_amount
        time.sleep(2)
        
        xpath_date_npi='//*[@id="ddeclm1500_txtNPI01"]'
        driver.find_element(By.XPATH, xpath_date_npi).click()
        
        xpath_units='//*[@id="ddeclm1500_txtDaysOrUnits01"]'
        driver.find_element(By.XPATH, xpath_units).click()
        driver.find_element(By.XPATH, xpath_units).send_keys(int(Total_units))   #Total_units
        time.sleep(2)
        
        
        xpath_NPI='//*[@id="ddeclm1500_txtProvidera"]'
        driver.find_element(By.XPATH, xpath_NPI).click()
        driver.find_element(By.XPATH, xpath_NPI).send_keys(NPI)   #NPI
        time.sleep(2)
        
        xpath_representative='//*[@id="ddeclm1500_txtPhysicianSignature3"]'
        driver.find_element(By.XPATH, xpath_representative).click()
        driver.find_element(By.XPATH, xpath_representative).send_keys("Praveena Kaliyaperumal")   #representative
        time.sleep(2)
        
        xpath_representative='//*[@id="ddeclm1500_txtPhysicianDate3"]'
        driver.find_element(By.XPATH, xpath_representative).click()
        driver.find_element(By.XPATH, xpath_representative).send_keys(formatted_time)   #formatted_time
        time.sleep(2)
        
        xpath_totalcharge='//*[@id="ddeclm1500_txtTotalCharge"]'
        driver.find_element(By.XPATH, xpath_totalcharge).click()
        driver.find_element(By.XPATH, xpath_totalcharge).send_keys("{:.2f}".format(total_BLL))   #total_BLL
        time.sleep(2)
        xpath_submit="//*[@id='ddeclm1500_btnSubmit']"
        driver.find_element(By.XPATH, xpath_submit).click()
        submitted_data=str(youth_last_name)+" "+str(youth_first_name) +" "+str(DOS_start_date_value)+" submitted"
        submitted_client.append(submitted_data)
        print(submitted_client)
        time.sleep(5)
        
        '''if total_BLL== bill_amount:
            print("its equal")
            xpath_submit="//*[@id='ddeclm1500_btnSubmit']"
            driver.find_element(By.XPATH, xpath_submit).click()
            time.sleep(5)
            # print(driver.find_element(By.XPATH,xpath_submit).is_displayed())
            # if Boolean_Display==False:
            #     print("successfully submitted")
            # else:
            #     print("report not submitted")
            # time.sleep(2)
        
        else:
            xapth_cancel="//*[@id='ddeclm1500_btnCancel']"
            driver.find_element(By.XPATH, xapth_cancel).click()
            print("value not matching, form not submitted")
        
        '''
        
        driver.switch_to.default_content()
        
        driver.close()
        
print(submitted_client)
print(not_submitted_client)

