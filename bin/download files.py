import wget
import requests
import csv
import os
##import win32com.client
from datetime import datetime

def ensure_dir(folder_path):
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

### Found here: http://stackoverflow.com/questions/15492436/how-to-create-a-url-shortcut-by-python
##def create_shortcut(url,shortcut_path):
##    ws = win32com.client.Dispatch("wscript.shell")
##    scut = ws.CreateShortcut(shortcut_path)
##    scut.TargetPath=url
##    scut.Save()

def get_url_list():
    url = requests.get("http://www.mssmallbiz.com/ericligman/Key_Shorts/MSFTFreeEbooks.txt")
    url_list = url.content.split("\r\n")
    url_list.pop(0)
    return url_list

startTime = datetime.now()
print startTime

#Setting up default Paths
INPUT_PATH = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.dirname(os.path.normpath(INPUT_PATH))
##INPUT_DIR = ensure_dir(os.path.join(BASE_DIR, "input"))
OUTPUT_DIR = ensure_dir(os.path.join(BASE_DIR, "output"))
##INPUT_CSV_FILE = os.path.join(INPUT_DIR, "input.csv")

#Reading in URL links from CSV
i = 1

url_list = get_url_list()
##with open(INPUT_CSV_FILE, 'rb') as csvfile:
##    reader = csv.DictReader(csvfile, dialect='excel',delimiter = ',')
for url_string in url_list:
    url = requests.get(url_string)
    filename = url.url.split("/")[-1]
    ext = filename.split(".")[-1].lower()
    if len(ext) > 5:
        folderpath = os.path.join(OUTPUT_DIR,"website")
        ensure_dir(folderpath)
        filepath = os.path.join(folderpath,filename+".url")
        print "Creating Internet Shortcut for link #{}...................Name: {}".format(i,filename)
        try:
            create_shortcut(url_string,filepath)
        except:
            print "Unable to create a shortcut for some reason. Moving on...."
        i+=1
    else:
        folderpath = os.path.join(OUTPUT_DIR,ext)
        ensure_dir(folderpath)
        filepath = os.path.join(folderpath,filename)
        print "Downloading link #{}...................Name: {}".format(i,filename)
        wget.download(url_string, filepath)
        i+=1

print "......................................................................End Runtime: ", datetime.now()-startTime