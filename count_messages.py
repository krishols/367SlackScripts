import numpy as np 
import pandas as pd
import sys
import re
import openpyxl

#sys.argv[1] is the csv file to be imported


in_file = sys.argv[1] 
threads = (pd.read_csv(in_file, delimiter = ","))
chat_ids = [] 
message_track = []
num_chat_ids = threads["thread"].iloc[-1] 
index = 1;
while index < num_chat_ids + 1:
    message_track.append(0)
    chat_ids.append(index)
    index += 1;
for row in threads.itertuples(): 
    message_track[row.thread-1] += 1
data = {"# messages": []}
df = pd.DataFrame(data)
df["# messages"] = message_track
#sys.argv[2] is the excel file to write to
#sys.argv[3] is the column to start dumping the dataframe
##writer.save()
writer = pd.ExcelWriter(sys.argv[2], engine = 'openpyxl')
book = openpyxl.load_workbook(sys.argv[2]) 
writer.book = book 
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
#sys.argv[2] is the excel file to write to
#sys.argv[3] is the column to start dumping the dataframe
df.to_excel(writer, index = False, startcol = int(sys.argv[3]))
writer.save()