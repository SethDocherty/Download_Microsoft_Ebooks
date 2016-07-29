import wget
import requests
import csv
import os
from datetime import datetime

def Default_Paths():
    INPUT_PATH = os.path.dirname(os.path.abspath(__file__))
    BASE_DIR = os.path.dirname(os.path.normpath(INPUT_PATH))
    OUTPUT_DIR = ensure_dir(os.path.join(BASE_DIR, "output"))
    return (INPUT_PATH, BASE_DIR, OUTPUT_DIR)


def ensure_dir(folder_path):
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    return folder_path
       
def get_url_list():
    url = requests.get("http://www.mssmallbiz.com/ericligman/Key_Shorts/MSFTFreeEbooks.txt")
    url_list = url.content.split("\r\n")
    url_list.pop(0)
    return url_list

## Found here: http://stackoverflow.com/questions/15492436/how-to-create-a-url-shortcut-by-python
#def create_shortcut(url,shortcut_path):
#    try:
#        import win32com.client
#        ws = win32com.client.Dispatch("wscript.shell")
#        scut = ws.CreateShortcut(shortcut_path)
#        scut.TargetPath=url
#        scut.Save()
#    except:
#        "Not Running in Pywin, Adding URL to text file"
#        temp = url_storage()[0]
#        temp.append(url)
#        url_storage(temp)

def url_storage(*values):
    url_storage.values = values or url_storage.values
    return url_storage.values

def export_to_csv_list(input_list, header, output_name):
    (INPUT_PATH, BASE_DIR, OUTPUT_DIR) = Default_Paths()
    Output_Path = os.path.join(OUTPUT_DIR,output_name + ".csv")
    writer = csv.writer(open(Output_Path, 'wb'))
    if type(header) is not type([]):
        writer.writerow([header])
    else:
        writer.writerow(header)
    for row in input_list:
        if type(row) is type([]):
            writer.writerow(row)
        else:
            row = [row]
            writer.writerow(row)

startTime = datetime.now()
print startTime

#Setting up default Paths
(INPUT_PATH, BASE_DIR, OUTPUT_DIR) = Default_Paths()

url_list = get_url_list()
totalLinkNum = len(url_list)
i = 1 # Counter

for url_string in url_list:
    url = requests.get(url_string)
    filename = url.url.split("/")[-1]
    ext = filename.split(".")[-1].lower()
    if len(ext) > 5:
        #folderpath = os.path.join(OUTPUT_DIR,"website")
        #ensure_dir(folderpath)
        url_link = url.url
        #filepath = os.path.join(folderpath,url_link+".url")
        print "Creating Internet Shortcut for link #{} of {}...................URL: {}".format(i,totalLinkNum,url_link)
        try:
            temp = url_storage()[0]
            temp.append(url_link)
            url_storage(temp)
        except:
            url_storage(url_link)
        #try:
        #    create_shortcut(url_string,filepath)
        #except:
        #    print "Unable to create a shortcut for some reason. Moving on...."
        i+=1
    else:
        folderpath = os.path.join(OUTPUT_DIR,ext)
        ensure_dir(folderpath)
        filepath = os.path.join(folderpath,filename)
        print "Downloading link #{} of {}...................File Name: {}".format(i,totalLinkNum,filename)
        wget.download(url_string, filepath)
        i+=1

folderpath = ensure_dir(os.path.join(OUTPUT_DIR,"website"))
export_to_csv_list(url_storage()[0], "Website Links", folderpath)

print "......................................................................End Runtime: ", datetime.now()-startTime