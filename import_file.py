import numpy as np 
import pandas as pd
import sys


in_file = sys.argv[1] 
threads = (pd.read_csv(in_file, delimiter = ","))
writer = pd.ExcelWriter("output.xlsx")
threads.to_excel(writer)
writer.save()
print(threads)