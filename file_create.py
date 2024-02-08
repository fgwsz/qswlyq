# -*- coding: utf-8 -*-
import os
import shutil
from datetime import datetime

def get_root_path():
    return os.path.dirname(os.path.abspath(__file__))

def get_current_date():
    date=datetime(datetime.now().year,datetime.now().month,datetime.now().day,0,0,0)
    return date.strftime("%Y%m%d")

def get_file_path():
    return get_root_path()+"\\"+get_current_date()+".docx"

if __name__=="__main__":
    if os.path.exists(get_file_path()):
        print(get_file_path()+" 已存在！")
    else:
        shutil.copy(get_root_path()+"\\file_template.docx",get_file_path())
        print(get_file_path()+" 创建成功！")
