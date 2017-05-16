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
xfile = openpyxl.load_workbook('jmu.xlsx')
# Open the worksheet we want to edit
sheet = xfile.get_sheet_by_name('students')
# Open the finished playing sound
pygame.init()
pygame.mixer.music.load('note.mp3')

# Some servers get annoyed if you make too many requests so dont do them all at once
# Start here
start = 2000
# End here
end = sheet.max_row + 1
# end = 2000

for row in range (start, end):
    # A. 
    lastName = sheet['C' + str(row)].value
    firstName = sheet['B' + str(row)].value
    # B.
    url = 'http://www.jmu.edu/cgi-bin/peoplestudentcms'
    post_fields = {'pattern': firstName + " " + lastName}
    request = Request(url, urlencode(post_fields).encode())
    json = urlopen(request).read()
    # print (json)
    # Make sure there are any results for the search
    if "<table id" in str(json):
        try:
            html = pandas.read_html(json)
            for i in range(0, len(html)):
                if firstName in str(html[i]) and lastName in str(html[i]):
                    p = re.compile('\w*@.*jmu\.edu')
                    m = p.search(str(html[i]))
                    if (m):
                        sheet['D' + str(row)] = m.group()
                        # Keep track of how close we are to being done
                        print (str(format((row - start) / (end - start) * 100.00, '.2f')) + "%: " + m.group())
        except Exception:
            pass
xfile.save('test.xlsx')
pygame.mixer.music.play()
time.sleep(3)
pygame.mixer.music.stop()
