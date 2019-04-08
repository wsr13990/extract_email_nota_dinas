# -*- coding: utf-8 -*-
"""
Created on Fri Feb 15 16:41:09 2019

@author: wahyu
"""

import os
import pandas as pd
import datetime
import win32com.client
import datetime as dt

FILE_REKAP = 'New Products Assmt.xlsx'
TXT_FILE = 'rekap_nodin.txt'
FOLDER='new_product'
PREV_DAY=100

class OutlookLib:
    def __init__(self, settings={}):
        self.settings = settings

    # Gets all messages in outlook   
    def get_messages(self,folder=None, n = 3):      
        outlook = win32com.client.Dispatch("Outlook.Application")

        # This allows us to access the "folder" hierarchy accessible within
        # Outlook. You can see this hierarchy yourself by opening Outlook
        # manually and bringing up the folder menu
        # (which typically says "Inbox" or "Outlook Today" or something).
        ns = outlook.GetNamespace("MAPI")
        
        if folder == None:
            messages = ns.GetDefaultFolder(6).Items
        else:
            messages = ns.GetDefaultFolder(6).Folders(FOLDER).Items
        messages.Sort("[ReceivedTime]", True)
        last_n_day = dt.datetime.now() - dt.timedelta(days = n)
        last_n_day = last_n_day.strftime('%m/%d/%Y %H:%M %p') 
        messages = messages.Restrict("[ReceivedTime] >= '" + last_n_day +"'")
        print(messages)
        return messages

    def get_body(self, msg):
        return msg.Body

    def get_subject(self, msg):
        subject = msg.Subject
        subject = subject.replace("FW: ","")
        subject = subject.replace("Fwd: ","")
        return subject
    
    def get_nodin(self, msg):
        nodin = msg.Subject
        try:
            nodin = nodin.split("(")[-1].split(")")[-2]
        except IndexError as e:
            print('Failed to get nodin number due to invalid character')
            nodin = 'NODIN_ERROR'
        return nodin

    def get_sender(self, msg):
        return msg.SenderName

    def get_recipient(self, msg):
        return msg.To

    def get_attachments(self, msg):
        return msg.Attachments


def Main():
    global attach

    outlook = OutlookLib()
    messages = outlook.get_messages(FOLDER,PREV_DAY)

    dir_path = os.path.dirname(os.path.realpath(__file__))
    # Loop all messages
    msg = messages.GetFirst()
    if msg == None:
        print('Failed to read email due to unknown error, try change PREV_DAY variable to larger value')
    while msg:
        rekap = pd.read_excel(FILE_REKAP,sheet_name='Sheet1')
        nodin = outlook.get_nodin(msg)
        subject = outlook.get_subject(msg)
        folder = outlook.get_nodin(msg).replace('/','').replace('.','')
        if len(rekap[rekap['Nota Dinas'].str.strip()==nodin]) >= 1:
            pass
        else:
            num = max(rekap['No'], default=0) + 1
            dest_path = os.path.join(dir_path,str(num)+'. '+folder)
            if os.path.isdir(dest_path) == False:
                os.mkdir(dest_path)
            if os.path.isdir(dest_path):
                attach = []
                if not len(msg.Attachments) is 0:
                    attach.append((msg.Attachments, msg.Subject))
                for attachTuple in attach:
                    print("Downloading attachments under " + attachTuple[1])
                    for fileAtt in attachTuple[0]:
                        print(dest_path+'/'+fileAtt.FileName)
                        fileAtt.SaveAsFile(dest_path+'/'+ fileAtt.FileName)
                rekap = rekap.append({'No' : num ,'Nota Dinas': nodin, 'Title / Subject' : subject,
                          'Tanggal Terima' : datetime.datetime.now(),
                          'Assessment Status' : 'not started'} , ignore_index=True)
        msg = messages.GetNext()
        print('Done downloading '+nodin)
        rekap.to_excel(FILE_REKAP,index=False)
        rekap['Folder'] = rekap['Nota Dinas'].map(lambda x: x.replace('.',''))
        rekap['Folder'] = rekap['Folder'].map(lambda x: x.replace('/',''))
        rekap[['Folder','Nota Dinas','Title / Subject','Tanggal Terima']].to_csv(TXT_FILE,sep='|', index=False, header=True)

if __name__ == "__main__":
    Main()
