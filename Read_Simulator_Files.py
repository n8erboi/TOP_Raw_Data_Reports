import time
import easygui
import os
import pandas as pd


#  Create timestamp for output filename
timestr = time.strftime("%Y_%m_%d-%H%M%S")

#  Raw data files filepath


raw_file_folder = easygui.diropenbox("Select The Directory Of Raw Data Files")
summary_file_folder = easygui.diropenbox("Select The Directory To Save The Summary File To")

oru_file_dict = {}
line_counter = 0

for path, subdirs, files in os.walk(raw_file_folder):
    for file in files:
        if 'd' in file.lower() and '.txt' in file:

            #oru_file_dict[file] = []

            #file_id = file.split('_')[0] + '_' + file.split('_')[1] + '_' + file.split('_')[2]

            filetype = 'utf-8'
            #  Check if file is utf-16-le
            try:
                for line in open(raw_file_folder + '\\' + file, encoding=filetype):
                    continue
            except UnicodeDecodeError:
                FileType = 'utf-16'


            #  Read each line of file
            if filetype == 'utf-8':

                read_counter = 0

                for line in open(raw_file_folder + '\\' + file, encoding=filetype):

                    #  Remove new line character at the end of the current line
                    line = line.rstrip("\n")

                    splitline = line.split(',')

                    am_value = splitline[0]

                    oru_file_dict[line_counter] = [file, read_counter, am_value]
                    read_counter += 1
                    line_counter += 1



raw_file_overall_df = pd.DataFrame.from_dict(oru_file_dict, "index")
raw_file_overall_df[2] = pd.to_numeric(raw_file_overall_df[2], errors='coerce')


raw_file_overall_pivot_df = raw_file_overall_df.pivot(index=1, columns=0, values=2)
raw_file_overall_pivot_df.sort_index()

with pd.ExcelWriter(summary_file_folder + "\\" +
                    'Sim-File-Summary_' +
                    timestr + ".xlsx", engine='xlsxwriter') as writer:
    raw_file_overall_pivot_df.to_excel(writer, sheet_name='Summary', index=True, header=True)



