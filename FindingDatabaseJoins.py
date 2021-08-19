# Imports
import glob
import pandas as pd
from pathlib import Path
from tkinter import filedialog
from tkinter import *
import os

# Setup and execute tkinter window
root = Tk()
root.withdraw()
folder_selected = filedialog.askdirectory()
output = os.path.join(folder_selected+'\output') #Setup the output folder based on tkinter input

files = glob.glob(folder_selected+'\*.sql')  # make a folder with the sql scripts of all stored procs
# make sure to leave \*.sql in order to grab all of the text files from the folder

if not os.path.exists(output): # If an 'output' folder does not exist already, then make it
    os.makedirs(output)

testoutput = os.path.join(output+r'\test.xlsx')  # set output file name

joins = []
for file in files:
    with open(file, 'r') as file_handle:
        filename = Path(file).stem
        filename = filename[4:-16]
        for line in file_handle:
            line = line.lstrip()
            line = line.lower()
            line = str(line)
            if line.find('between') != -1:
                pass
            elif line.find('@') != -1:
                pass
            elif line.find('not') != -1:
                pass
            elif line.find('<') != -1:
                pass
            elif line.find('>') != -1:
                pass
            elif line.find('like') != -1:
                pass
            elif line.find(r"'") != -1:
                pass
            elif line.startswith('from'):
                joins.append({'join':line, 'file':filename})
            elif line.startswith('inner'):
                joins.append({'join':line, 'file':filename})
            elif line.startswith('left'):
                joins.append({'join':line, 'file':filename})
            elif line.startswith('on'):
                joins.append({'join': line, 'file':filename})
            elif line.startswith('and'):
                joins.append({'join': line, 'file':filename})
            else:
                pass

joins = pd.DataFrame(joins)


df = joins
print(df)


df = (df.groupby((~df['join'].str.startswith('on ')).cumsum())
   .agg(' '.join)
   .reset_index(drop=True))

df = (df.groupby((~df['join'].str.startswith('and ')).cumsum())
   .agg(' '.join)
   .reset_index(drop=True))

df['file'] = df['file'].str.split(' ').str[0]


print(df)
df.to_excel(testoutput)
