# This file loops through all txt files in a folder and write out every instance of a keyword. In this case "XXXX" and "XXXX.["
# Next to the keyword, it also outputs the name of the file (line 15 will need to be updated)
# It then sends these files to excel for viewing

import glob
files = glob.glob(r'XXXXXXXXXXXXXXXXXXXXXXXXXXXX\*.txt') # make a folder with the txt scripts of all stored procs
# make sure to leave \*.txt in order to grab all of the text files from the folder

tidyoutput = r'xxxxxxxxxxxxxxxxxxx\wordstidy.xlsx' # set output file name
messyoutput =r'xxxxxxxxxxxxxxxxxxx\finaltest.xlsx' # set output file name
import pandas as pd

collect = []
for file in files:
    with open(file, 'r') as file_handle:
        filename = file[61:-20]
        for line in file_handle:
            for word in line.split():
                if word.startswith('FAST'):
                    collect.append({'keyword': word, 'filename': filename})
                elif word.startswith('FAST.['):
                    collect.append({'keyword': word, 'filename': filename})
                elif word.startswith('Fast'):
                    collect.append({'keyword': word, 'filename': filename})
                elif word.startswith('Fast.['):
                    collect.append({'keyword': word, 'filename': filename})
                elif word.startswith('fast'):
                    collect.append({'keyword': word, 'filename': filename})
                elif word.startswith('fast.['):
                    collect.append({'keyword': word, 'filename': filename})
                else:
                    pass
                
words_files_tidy = pd.DataFrame.from_records(collect).drop_duplicates()
words_files_tidy.to_excel(tidyoutput)
final_df = words_files_tidy.pivot(index='keyword', columns='filename', values='filename').reset_index()
final_df.to_excel(messyoutput)
