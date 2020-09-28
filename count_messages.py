import numpy as np 
import pandas as pd
import sys
import re
import openpyxl

#sys.argv[1] is the csv file to be imported
#creates DataFrame from csv file
in_file = sys.argv[1] 
threads = (pd.read_csv(in_file, delimiter = ","))
chat_ids = [] 
message_track = []
num_chat_ids = threads["thread"].iloc[-1] 
index = 1;

while index < num_chat_ids + 1:
    #creates list message_track with index for each thread
    #creates list chat_ids to keep track of each thread 
    message_track.append(0)
    chat_ids.append(index)
    index += 1;
    
for row in threads.itertuples(): 
    #checks row for the thread and adds one to corresponding thread to represent a message 
    message_track[row.thread-1] += 1
data = {"# messages": []}

#creates DataFrame of number of messages per thread
df = pd.DataFrame(data)
df["# messages"] = message_track

#sys.argv[2] is the excel file to update
#uploads existing excel file to update
writer = pd.ExcelWriter(sys.argv[2], engine = 'openpyxl')
book = openpyxl.load_workbook(sys.argv[2]) 
writer.book = book 
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

#sys.argv[2] is the excel file to write to
#sys.argv[3] is the column to start dumping the dataframe
#updates excel sheet with created DataFrame
df.to_excel(writer, index = False, startcol = int(sys.argv[3]))
writer.save()