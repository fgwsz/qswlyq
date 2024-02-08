# -*- coding: utf-8 -*-
import os
from datetime import datetime
import subprocess

def get_root_path():
    return os.path.dirname(os.path.abspath(__file__))

def get_current_date():
    date=datetime(datetime.now().year,datetime.now().month,datetime.now().day,0,0,0)
    return date.strftime("%Y%m%d")

def get_file_path():
    return get_root_path()+"\\"+get_current_date()+".docx"

if __name__=="__main__":
    docx_path=get_file_path()
    command=f'explorer.exe "{docx_path}"'
    subprocess.run(command,shell=True)
