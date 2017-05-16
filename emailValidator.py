from urllib.parse import urlencode
from urllib.request import Request, urlopen
from validate_email import validate_email
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import py3dns
import validate_email
import pandas
import openpyxl
import re
import math
 
xfile = openpyxl.load_workbook('UoR.xlsx')
sheet = xfile.get_sheet_by_name('students')
 
start = 2
end = 3
for row in range(start, end):
    #email = sheet['D' + str(row)].value
    email = 'twfischl@gmail.com'
    print(str(row)+ " validating " + email)
    is_valid = validate_email(email, verify=True)
    print(is_valid)
    if is_valid == False:
        sheet['D' + str(row)].value = None
        print("invalidated " + email)
    if is_valid == True:
        print("valid")
xfile.save('test.xlsx')
