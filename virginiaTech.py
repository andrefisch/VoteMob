import re
import smtplib
import dns.resolver
import socket
import openpyxl

xfile = openpyxl.load_workbook('UoR.xlsx')
sheet = xfile.get_sheet_by_name('students')

start = 2
end = 3
fromAddress = 'election.virginiatech@gmail.com'
regex = '^[_a-z0-9-]+(\.[_a-z0-9-]+)*@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,4})$'
for row in range(start, end):
    #addressToVerify = sheet['D' + str(row)].value
    addressToVerify = 'twfischl@gmail.com'
    match = re.match(regex, addressToVerify)
    if match == None:
        print('Bad Syntax')
        sheet['D' + str(row)].value = None
    else:
        splitAddress = addressToVerify.split('@')
        domain = str(splitAddress[1])
        print('Domain:', domain)
        records = dns.resolver.query(domain, 'MX')
        mxRecord = records[0].exchange
        mxRecord = str(mxRecord)
        host = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        server = smtplib.SMTP()
        server.set_debuglevel(0)
        print (mxRecord)
        server.connect(mxRecord, 587)
        server.helo(server.local_hostname) ### server.local_hostname(Get local server hostname)
        server.mail(fromAddress)
        code, message = server.rcpt(str(addressToVerify))
        server.quit()
        host.close()
        if code == 250:
            print(str(addressToVerify) + " is good.")
            pass
        else:
            print(str(addressToVerify) + ' is bad.')
            sheet['D' + str(row)].value = None
