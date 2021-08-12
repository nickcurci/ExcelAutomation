# This file loops through all sql files in a folder and write out every instance of a keyword (or multiple)
# It then sends these files to excel for viewing

import glob
import pandas as pd
from pathlib import Path
from tkinter import filedialog
from tkinter import *
import os
root = Tk()
root.withdraw()
folder_selected = filedialog.askdirectory()
output = os.path.join(folder_selected+'\output')

files = glob.glob(folder_selected+'\*.sql') # make a folder with the sql scripts of all stored procs
# make sure to leave \*.sql in order to grab all of the text files from the folder
if not os.path.exists(output):
    os.makedirs(output)
tidyoutput = os.path.join(output+'\wordstidy.xlsx') # set output file name
messyoutput =os.path.join(output+r'\finaltest.xlsx') # set output file name


collect = []
for file in files:
    with open(file, 'r') as file_handle:
        filename = Path(file).stem
        filename = filename[4:-16]
        for line in file_handle:
            for word in line.split():
                word = word.replace('[','')
                word = word.replace(']','')
                word = word.lower()
                if word.count('.') > 2:
                    pass
                elif word.startswith('FAST_'):
                    pass
                elif word.startswith('Fast_'):
                    pass
                elif word.startswith('FASTQ'):
                    pass
                elif word.startswith('FastQ'):
                    pass
                elif word.endswith(','):
                    pass
                elif word.endswith(')'):
                    pass
                elif word.endswith('.'):
                    pass
                elif word.startswith('FAST.'):
                    collect.append({'table': word, 'storedproc': filename})
                elif word.startswith('FAST.['):
                    collect.append({'table': word, 'storedproc': filename})
                elif word.startswith('Fast.'):
                    collect.append({'table': word, 'storedproc': filename})
                elif word.startswith('Fast.['):
                    collect.append({'table': word, 'storedproc': filename})
                elif word.startswith('fast.'):
                    collect.append({'table': word, 'storedproc': filename})
                elif word.startswith('fast.['):
                    collect.append({'table': word, 'storedproc': filename})
                else:
                    pass

words_files_tidy = pd.DataFrame.from_records(collect).drop_duplicates()
words_files_tidy.to_excel(tidyoutput, index = False)
final_df = words_files_tidy.pivot(index='table', columns='storedproc', values='storedproc').reset_index()
final_df.to_excel(messyoutput)
