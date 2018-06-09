import win32com.client
import sys
import unicodecsv as csv
import re
import datetime
import pymsgbox
import docx
from __future__ import print_function
from mailmerge import MailMerge
from datetime import date

Ex_list = []
Ex_extract = False
loop_run = False

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
                                    # the inbox. 
messages = inbox.Items.Restrict("[ReceivedTime] >= '" + str(datetime.date.today()) + "'")

for message in reversed(messages):
    if loop_run:
        break
    subject = message.subject
    Ex_extract = False
    if re.search("Mail Subject", str(subject)) :      
        body = message.Body
        body_list = body.splitlines()
        
        for val in body_list:
            if Ex_extract:
                Ex_list.append(val)
            if re.search("Example :", val):
                Ex_extract = True
                
        loop_run = True
        
template = 'ABC_Template.docx'
document = MailMerge(template)
print(document.get_merge_fields())

document.merge(
    Example = ', '.join(Ex_list)
    )
document.write('ABC.docx')
document.close()