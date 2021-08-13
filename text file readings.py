# This file loops through all sql files in a folder and write out every instance of a keyword (or multiple)
# It then sends these files to excel for viewing
# The commented portion sends an email with the files as well

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
tidyoutput = os.path.join(output+r'\tables_tidy.xlsx') # set output file name
messyoutput =os.path.join(output+r'\tables_messy.xlsx') # set output file name


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


# import win32com.client
#
# outlook = win32com.client.Dispatch('outlook.application')
#
# mail = outlook.CreateItem(0)
#
# mail.To = 'xxxx@xxxx.com'
# mail.Subject = 'Finding Fast Files'
# mail.HTMLBody = (
#                 '<p>Good Morning Chris!</p>'
#                 '<p>This is an email sent stright from python, with attachments that are created and exported to excel stright from the code.</p>'
#                 '<p>Pretty neat. Enjoy the rest of your trip!</p>'
#                 '<p> </p>'
#                 '<p> </p>'
#                 '<strong>Name</strong> | Associate Business Intelligence Analyst</p>'
#                 'Amica Life Insurance Company | Life - Data Services</p>'
#                 '10 Amica Center Blvd. | Lincoln, RI | 02865</p>'
#                 'Voice: xxx-xxx-xxxx <br /> <a href="mailto:ncurci@amica.com">xxxx@xxxxx.com</a> | <a href="blocked::http://www.amica.com/">Amica.com</a>'
# )
# # mail.Body = "Test."
# mail.Attachments.Add(tidyoutput)
# mail.Attachments.Add(messyoutput)
# mail.CC = 'xxxxx@xxxxx.com'
# mail.Send()
