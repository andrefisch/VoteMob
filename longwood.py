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
xfile = openpyxl.load_workbook('longwood.xlsx')
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
# end = 100

for row in range (start, end):
    # A. 
    firstName = sheet['B' + str(row)].value
    lastName = sheet['C' + str(row)].value
    firstName = firstName.replace(" ", "+")
    lastName = lastName.replace(" ", "+")
    # print(str(row)+ " searching for " + firstName + " " + lastName)
    # B.
    url = 'http://www.longwood.edu/directory/?query=' + firstName + "+" + lastName
    post_fields = {'query': firstName + " " + lastName}
    request = Request(url)#, urlencode(post_fields).encode())
    json = urlopen(request).read()
    #print (json)
    # Make sure there are any results for the search
    if "<table" in str(json):
        try:
            html = pandas.read_html(json)
            for i in range(1, len(html[0][0])):
                if firstName in str(html[0][0][i]) and lastName in str(html[0][0][i]):
                    # p = re.compile('\w*@.*longwood\.edu')
                    # m = p.search(str(html[i]))
                    # print(html[0][1][i])
                    sheet['D' + str(row)] = html[0][1][i]
                    print (str(format((row - start) / (end - start) * 100.00, '.2f')) + "%: " + html[0][1][i])
                    break
                    # Keep track of how close we are to being done
        except Exception:
            pass
xfile.save('test.xlsx')
pygame.mixer.music.play()
time.sleep(3)
pygame.mixer.music.stop()
