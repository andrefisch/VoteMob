from urllib.parse import urlencode
from urllib.request import Request, urlopen
import pandas
import openpyxl
import re
import math
import pygame, time

'''
1. import spreadsheet
2. for loop:
    A. find value in cell C(row)
    B. make a request to database
    C. find student in response using last name C(row) and first name B(row)
    D. if name exists in database:
        a. replace empty cell D(row) with email address
3. save file
'''

# 1.
# Open the file for editing
xfile = openpyxl.load_workbook('umw.xlsx')
# Open the worksheet we want to edit
sheet = xfile.get_sheet_by_name('students')
# Open the finished playing sound
pygame.init()
pygame.mixer.music.load('note.mp3')

# Some servers get annoyed if you make too many requests so dont do them all at once
# Start here
start = 2
# End here
end = sheet.max_row + 1
# end = 2000

for row in range (start, end):
    # A. 
    firstName = sheet['B' + str(row)].value
    lastName = sheet['C' + str(row)].value
    # B.
    url = 'http://students.umw.edu/directory/'
    post_fields = {'adeq': firstName + " " + lastName}
    request = Request(url, urlencode(post_fields).encode())
    json = urlopen(request).read()
    #print (json)
    # Make sure there are any results for the search
    #print (html[1])
    p = re.compile('\w*@mail\.umw\.edu')
    m = p.search(str(json))
    if (m):
        sheet['D' + str(row)] = m.group()
        # Keep track of how close we are to being done
        print (str(format((row - start) / (end - start) * 100.00, '.2f')) + "%: " + m.group())
xfile.save('test.xlsx')
pygame.mixer.music.play()
time.sleep(3)
pygame.mixer.music.stop()
