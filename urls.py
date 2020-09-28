import numpy as np 
import pandas as pd
import sys
import re

#meant to be the first script ran to create a new excel file 
#sys.argv[1] is the csv file to be imported and create a dataframe for it
in_file = sys.argv[1] 
threads = (pd.read_csv(in_file, delimiter = ","))
chat_ids = [] 
url_track = []
num_chat_ids = threads["thread"].iloc[-1] 
index = 1;

while index < num_chat_ids + 1:
    #creates list url_track with index for each thread
    #creates list chat_ids to keep track of each thread 
    url_track.append(0)
    chat_ids.append(index)
    index += 1;
    
for row in threads.itertuples(): 
    #checks each row message for a url using regular expression operation
    #adds number of urls found to thread index in url_track
    url_sum = 0
    regex = r"(?i)\b((?:https?://|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}/)(?:[^\s()<>]+|\(([^\s()<>]+|(\([^\s()<>]+\)))*\))+(?:\(([^\s()<>]+|(\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:'\".,<>?«»“”‘’]))"
    url_sum += len(re.findall(regex,row.message))
    url_track[row.thread -1] += url_sum
data = {"chat_id": [], "# of URLs": []}

#creates DataFrame of threads and # of urls per thread
df = pd.DataFrame(data)
df["chat_id"] = chat_ids
df["# of URLs"] = url_track

#sys.argv[2] is the excel file to create
#sys.argv[3] is the column to start dumping the dataframe
df.to_excel(sys.argv[2], index = False, startcol = int(sys.argv[3]))
