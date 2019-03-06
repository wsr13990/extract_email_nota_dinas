# -*- coding: utf-8 -*-
"""
Created on Fri Feb 15 16:41:09 2019

@author: wahyu
"""

import os
import shutil
from email import policy
from email.parser import BytesParser
import pandas as pd
import datetime

FILE_REKAP = 'New Products Assmt.xlsx'
TXT_FILE = 'rekap_nodin.txt'
MOVE_EML = True

files = [f for f in os.listdir('.') if os.path.isfile(f)]

def get_eml_body(eml_file):
    with open(eml_file, 'rb') as fp:
        msg = BytesParser(policy=policy.default).parse(fp)
        text = msg.get_body(preferencelist=('plain')).get_content()
        return text

def get_eml_subject(eml_file):
    with open(eml_file, 'rb') as fp:
        msg = BytesParser(policy=policy.default).parse(fp)
        return msg.get_all('Subject')[0]

def get_eml_nota_dinas(eml_file):
    return get_eml_subject(eml_file).split("(")[-1].split(")")[-2]

def download_eml_attachment(eml_file,dest_path,move_eml=True):
    is_success = False
    attachment_counter = 0
    dir_path = os.path.dirname(os.path.realpath(__file__))
    with open(eml_file, 'rb') as fp:
        msg = BytesParser(policy=policy.default).parse(fp)
        count_attachments = len(msg.get_payload())
        if count_attachments > 0 :
            for i in range(1,len(msg.get_payload())):
                attachment = msg.get_payload()[i]
                attachment_name = attachment.get_filename()
                try:
                    open(attachment_name,'wb').write(attachment.get_payload(decode=True))
                    attachment_name_old=attachment_name
                    attachment_name = attachment_name.replace(",","")
                    attachment_name = attachment_name.replace("&","")
                    attachment_name = attachment_name.replace(";","")
                    ext = attachment_name.split(".")[-1]
                    if (len(attachment_name) >= 60):
                        attachment_name = attachment_name[:60]+"."+ext
                        os.rename(attachment_name_old, attachment_name)
                    if os.path.isdir(dest_path):
                        shutil.move(os.path.join(dir_path, attachment_name),
                                    os.path.join(dest_path, attachment_name))
                        attachment_counter += 1
                except FileNotFoundError:
                    print('Possibly invalid character in the attachment file')
                print(os.path.isfile(os.path.join(dir_path, attachment_name)))
        fp.close()
        if count_attachments > 0 :
            if attachment_counter == count_attachments-1:
                is_success = True
            if is_success and move_eml == True:
                print('All attachment succesfully downloaded')
            else:
                print('Failed to download some attachment')
            if (len(eml_file) >= 60):
                        eml_file_new = eml_file[:60]+"."+'eml'
                        os.rename(eml_file, eml_file_new)
            shutil.move(os.path.join(dir_path, eml_file_new), os.path.join(dest_path, eml_file_new))

def main_eml():
    dir_path = os.path.dirname(os.path.realpath(__file__))
    rekap = pd.read_excel(FILE_REKAP,sheet_name='Sheet1')
    for eml_file in files:
        if eml_file.endswith('.eml'):
            subject = get_eml_subject(eml_file)
            subject = subject.replace("FW: ","")
            subject = subject.replace("Fwd: ","")
            nodin = get_eml_nota_dinas(eml_file)
            if len(rekap[rekap['Nota Dinas'].str.strip()==nodin]) >= 1:
                os.remove(eml_file)
                pass
            else:         
                num = max(rekap['No'], default=0) + 1
                folder = get_eml_nota_dinas(eml_file).replace('/','').replace('.','')
                dest_path = os.path.join(dir_path,str(num)+'. '+folder)
                if os.path.isdir(dest_path) == False:
                    os.mkdir(dest_path)
                if os.path.isdir(dest_path):
                    if MOVE_EML == True:
                        download_eml_attachment(eml_file,dest_path,True)
                    else:
                        download_eml_attachment(eml_file,dest_path,False)
                rekap = rekap.append({'No' : num ,'Nota Dinas': nodin, 'Title / Subject' : subject,
                                      'Tanggal Terima' : datetime.datetime.now(),
                                      'Assessment Status' : 'not started'} , ignore_index=True)
            rekap.to_excel(FILE_REKAP,index=False)
            rekap['Folder'] = rekap['Nota Dinas'].map(lambda x: x.replace('/',''))
            rekap[['Folder','Nota Dinas','Title / Subject','Tanggal Terima']].to_csv(TXT_FILE,sep='|', index=False, header=True)

            
            
##TODO: Add .msg functionality
                
main_eml()