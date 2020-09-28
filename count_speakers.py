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
speaker_track = []
num_chat_ids = threads["thread"].iloc[-1] 
index = 1;
while index < num_chat_ids + 1:
    #creates list speaker_track with index of a list for each thread
    #creates list chat_ids to keep track of each thread 
    speaker_track.append([])
    chat_ids.append(index)
    index += 1;
    
for row in threads.itertuples():
    #checks each row for speaker of the row
    #if speaker is unique to the thread's list of speaker_track, appends list with the speaker
    if row.speaker in speaker_track[row.thread-1]:
        pass
    else: 
        speaker_track[row.thread-1].append(row.speaker)
index = 0

while index < num_chat_ids: 
    #updates speaker_track list elements with the length of each list to represent number of unique speakers
    speaker_track[index] = len(speaker_track[index])
    index += 1
    
#creates DataFrame with number of speakers per thread 
data = {"# unique speakers": []}
df = pd.DataFrame(data)
df["# unique speakers"] = speaker_track

#sys.argv[2] is the excel file to write to
#uploads existing excel file to update 
writer = pd.ExcelWriter(sys.argv[2], engine = 'openpyxl')
book = openpyxl.load_workbook(sys.argv[2]) 
writer.book = book 
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

#sys.argv[3] is the column to start dumping the dataframe
#updates excel sheet with df DataFrame
df.to_excel(writer, index = False, startcol = int(sys.argv[3]))
writer.save()