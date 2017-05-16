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
xfile = openpyxl.load_workbook('vt.xlsx')
# Open the worksheet we want to edit
sheet = xfile.get_sheet_by_name('students')
# Open the finished playing sound
pygame.init()
pygame.mixer.music.load('note.mp3')

# Some servers get annoyed if you make too many requests so dont do them all at once
# Start here
start = 8000
# End here
end = sheet.max_row + 1
# end = 8000

for row in range (start, end):
    # A. 
    lastName = sheet['C' + str(row)].value
    # B.
    url = 'https://search.vt.edu/search/people.html'
    post_fields = {'q': lastName}
    request = Request(url, urlencode(post_fields).encode())
    json = urlopen(request).read()
    # Make sure there are any results for the search
    if "<table>" in str(json):
        firstName = sheet['B' + str(row)].value
        html = pandas.read_html(json)
        for i in range(0, len(html[0][0])):
            if lastName in html[0][0][i] and firstName in html[0][0][i]:
                p = re.compile('\w*@vt\.edu')
                # print (isinstance(html[0][2][i], str))
                if (isinstance(html[0][2][i], str)):
                    m = p.search(html[0][2][i])
                    if (m):
                        sheet['D' + str(row)] = m.group()
                        # Keep track of how close we are to being done
                        print (str(format((row - start) / (end - start) * 100.00, '.2f')) + "%: " + m.group())
xfile.save('test.xlsx')
pygame.mixer.music.play()
time.sleep(3)
pygame.mixer.music.stop()
