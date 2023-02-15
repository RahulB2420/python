import pandas as pd
import urllib.request
import regex as re
import time
import os


drive_path = r"local\drive\location" # Give path for drive location to upload
gsheet_path = r"local\excel\location" # Give path for excel file along with filename

gsheet = pd.read_excel(gsheet_path)

for row_num,row_val in gsheet.iterrows():
    m_run = -1 #adjust value for manual run or keep default value as -1 (Ex : If code errored out after 100 rows so after fixing error you want to run code from 101 row so change m_run value to 100)
    if row_num <= m_run: 
        pass
    else:       
        uname_tmp = row_val['User_email'].split(".")
        uname = uname_tmp[0] + " " + uname_tmp[1].split("@")[0]
        file_name = str(row_val['Video_name'])
        file_name = re.sub(r"[^a-zA-Z0-9]+", ' ', file_name)
        file_name = file_name.strip()
        id = str(row_num + 1)
        if len(file_name) == 0:
            file_name = id + "no_title"
        else:
            file_name = id + "-" + file_name + " - " + (str(row_val['Created_date']).split(":")[0]).replace("/","-") + " " + str(row_val['Created_date']).split(":")[1] + ".webm"
        upload_folder = drive_path + "\\" + uname 
        upload_file = str(upload_folder + "\\" + file_name)
        if os.path.exists(upload_folder):
            pass
        else:
            os.makedirs(upload_folder)
        try:
            os.chdir(upload_folder)
            cur_files = os.listdir()
            if file_name in cur_files:
                print(row_num + 1," Row"," skipped because the file is already present")
                continue
            else:
                urllib.request.urlretrieve(row_val['download_link'], upload_file)
                print(row_num + 1," Row(s)"," done")
            time.sleep(3)
            cur_files = os.listdir()
            if file_name not in cur_files:
                print("file ",file_name," on row ",row_num + 1," was not downloaded successfully")
            
        except Exception as e:
            print(upload_folder)
            print("Error on Row ",row_num + 1) 
            print(repr(e))
           

