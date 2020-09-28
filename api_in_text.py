import numpy as np 
import pandas as pd
import sys
import re
import openpyxl

in_file = sys.argv[1] 
threads = (pd.read_csv(in_file, delimiter = ","))
chat_ids = [] 
index = -1
api_track = []
list_of_api = ["wxPython", "PYGObject","Pmw", "WCK", "Tix", "MySQLdb", "PyGreSQL", "Gadfly", "SQLAlchemy", "KInterbasDB", "beautiful soup", "scrape", "mechanize", "libgmail", "google maps", "requests", "selenium", "pyquery", "python imaging library", "gdmodule", "videocapture", "moviepy", "pyscreenshot", "scipy", "matplotlib", "pandas", "numpy", "pygame", "pyglet", "pyOpenGL", "pysonic", "pymedia", "pmidi", "mutagen", "pywin32", "pyrtf", "wmi", "py2exe", "py2app", "pyobjc", "pyusb", "pyserial", "uspp", "twisted", "jabberpy", "pyexpect", "vpython"]
num_chat_ids = threads["thread"].iloc[-1] 
index = 1;
while index < num_chat_ids + 1:
    api_track.append(0)
    chat_ids.append(index)
    index += 1;
for row in threads.itertuples(): 
    api_sum = 0
    regex = r'(\w+[.]\w+)@'  
    api_sum += len(re.findall(regex,row.message))
    for api in list_of_api: 
        if api in row.message: 
            api_sum += 1; 
    api_track[row.thread-1] += api_sum
data = {"# API references": []}
df = pd.DataFrame(data)
df["# API references"] = api_track

writer = pd.ExcelWriter(sys.argv[2], engine = 'openpyxl')
book = openpyxl.load_workbook(sys.argv[2]) 
writer.book = book 
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
#sys.argv[2] is the excel file to write to
#sys.argv[3] is the column to start dumping the dataframe
df.to_excel(writer, index = False, startcol = int(sys.argv[3]))
writer.save()