from urllib.parse import urlencode
from urllib.request import Request, urlopen
import pandas
import openpyxl

'''
1. import spreadsheet
2. for loop:
    A. find value in cell M(row)
    B. make a request to W&M database
    C. find student in W&M response using first name M(row) and last name N(row)
    D. if name exists in database:
        a. replace empty cell T(row) with email address
3. save file
'''

# 1.
# Open the file for editing
xfile = openpyxl.load_workbook('williamAndMary.xlsx')
# Open the worksheet we want to edit
sheet = xfile.get_sheet_by_name('students')

# Start here
start = 9000
# End here
end = sheet.max_row + 1
# 2.
for row in range (start, end):
    # A. 
    lastName = sheet['M' + str(row)].value
    # B.
    url = 'http://directory.wm.edu/people/namelisting.cfm'
    post_fields = {'searchtype': 'last', 'criteria': 'same', 'phrase': lastName}
    request = Request(url, urlencode(post_fields).encode())
    json = urlopen(request).read()
    # Make sure there are any results for the search
    if "table class" in str(json):
        firstName = sheet['N' + str(row)].value
        html = pandas.read_html(json)
        for i in range(0, len(html[-1][0])):
            # C. && D.
            if (lastName.upper() + ", " + firstName) in html[-1][0][i]:
                # a.
                sheet['T' + str(row)] = html[-1][1][i]
                # Keep track of how close we are to being done
                print (str(format((row - start) / (end - start) * 100.00, '.2f')) + "%: " + html[-1][1][i])

xfile.save('test.xlsx')
