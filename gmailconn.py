#!/usr/bin/env python
#
# Very basic example of using Python 3 and IMAP to iterate over emails in a
# gmail folder/label.  This code is released into the public domain.
#
# This script is example code from this blog post:
#   
#
# This is an updated version of the original -- modified to work with Python 3.4.
#
import sys
import imaplib
import getpass
import email
import email.header
import datetime
import re


EMAIL_ACCOUNT = "juanjose.perez@globant.com"

# Use 'INBOX' to read inbox.  Note that whatever folder is specified, 
# after successfully running this script all emails in that folder 
# will be marked as read.
EMAIL_FOLDER = "CloudStudio/MB/Pingdom"
out = open("email_Report.csv","w")
print("to;sender;subject;messageid;date",file=out)


def process_mailbox(M):
    """
    Do something with emails messages in the folder.  
    For the sake of this example, print some headers.
    """

    rv, data = M.search(None, "ALL")
    if rv != 'OK':
        print("No messages found!")
        return

    for num in data[0].split():
        rv, data = M.fetch(num, '(RFC822)')
        if rv != 'OK':
            print("ERROR getting message", num)
            return
        
      #  for response_part in data:
      #      if instance(response_part,tuple):
      #          mesg = email.message_from_string(response_part[1])

        msg = email.message_from_bytes(data[0][1])
        hdr = email.header.make_header(email.header.decode_header(msg['Subject']))
        hdrfrom = email.header.make_header(email.header.decode_header(msg['From']))
        hdrto = email.header.make_header(email.header.decode_header(msg['To']))
        # hdrbody = email.header.make_header(email.header.decode_header(msg))
        #b = email.message_from_string(data[0][1])
        subject = re.sub('\n|\r', '', str(hdr))
        hfrom = re.sub('\n|\r', '', str(hdrfrom))
        hto = re.sub('\n|\r', '', str(hdrto))
        
        print('\n')
        print('RegistroNuevo')
        print('--------------------------------')
        #print('\n')
        #print('NuevaInformacion')
        print('Message %s: %s' % (num, subject))
        print('Raw Date:', msg['Date'])
        print('-------------HDRSubject---------------')
        #print(hdr)
        print(subject)
        #print() 
        print('-------------HDRfrom---------------')
        print(hfrom)
        #print(hdrfrom)
        print('-------------HDRTo---------------')
        print(hto)
        #print(hdrto)

        
        print("{0};{1};{2};{3};{4}".format(hto,hfrom,subject,msg['Message-ID'],msg['Date']),file=out)
        print("{0};{1};{2};{3};{4}".format(hto,hfrom,subject,msg['Message-ID'],msg['Date']))
        print('--------------------------------')
        #print('++++++++++++++++++++++++++++++++')
        #print(msg,file=out)
        #print('++++++++++++++++++++++++++++++++')
        
               # print('Raw Date:', msg['Date'])
       # print(b['body'])
        # Now convert to local date-time
        date_tuple = email.utils.parsedate_tz(msg['Date'])
        if date_tuple:
            local_date = datetime.datetime.fromtimestamp(
                email.utils.mktime_tz(date_tuple))
            print ("Local Date:", \
                local_date.strftime("%a, %d %b %Y %H:%M:%S"))


M = imaplib.IMAP4_SSL('imap.gmail.com')

try:
    rv, data = M.login(EMAIL_ACCOUNT, getpass.getpass())
except imaplib.IMAP4.error:
    print ("LOGIN FAILED!!! ")
    sys.exit(1)

print(rv, data)

rv, mailboxes = M.list()
if rv == 'OK':
    print("Mailboxes:")
    print(mailboxes)

rv, data = M.select(EMAIL_FOLDER)
if rv == 'OK':
    print("Processing mailbox...\n")
    process_mailbox(M)
    M.close()
else:
    print("ERROR: Unable to open mailbox ", rv)

M.logout()