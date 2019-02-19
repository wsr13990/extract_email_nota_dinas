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
MOVE_EML = False

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
                open(attachment_name,'wb').write(attachment.get_payload(decode=True))                
                if os.path.isdir(dest_path):
                    shutil.move(os.path.join(dir_path, attachment_name),
                                os.path.join(dest_path, attachment_name))
                    attachment_counter += 1
        fp.close()
        if count_attachments > 0 :
            if attachment_counter == count_attachments-1:
                is_success = True
            if is_success and move_eml == True:
                shutil.move(os.path.join(dir_path, eml_file), os.path.join(dest_path, eml_file))
                  
def main_eml():
    dir_path = os.path.dirname(os.path.realpath(__file__))
    rekap = pd.read_excel(FILE_REKAP,sheet_name='Sheet1')
    for eml_file in files:
        if eml_file.endswith('.eml'):
            subject = get_eml_subject(eml_file)
            nodin = get_eml_nota_dinas(eml_file)
            if len(rekap[rekap['Nota Dinas'].str.strip()==nodin]) >= 1:
                os.remove(eml_file)
            else:         
                num = max(rekap['No'], default=0) + 1
                folder = get_eml_nota_dinas(eml_file).replace('/','').replace('.','')
                dest_path = os.path.join(dir_path,str(num)+'. '+folder)
                if os.path.isdir(dest_path) == False:
                    os.mkdir(dest_path)
                if os.path.isdir(dest_path):
                    if MOVE_EML == True:
                        download_eml_attachment(eml_file,dest_path,False)
                    else:
                        download_eml_attachment(eml_file,dest_path)
                rekap = rekap.append({'No' : num ,'Nota Dinas': nodin, 'Title / Subject' : subject,
                                      'Tanggal Terima' : datetime.datetime.now(),
                                      'Assessment Status' : 'not started'} , ignore_index=True)
            rekap.to_excel(FILE_REKAP,index=False)
            
##TODO: Add .msg functionality
                
main_eml()