import time
import easygui
import os
import pandas as pd
import numpy as np
import io
from datetime import datetime

#  This script compiles ACL TOP raw RLU files into a single excel file and checks for any
#  assays that were run concurrently

#  Create timestamp for output filename
timestr = time.strftime("%Y_%m_%d-%H%M%S")

#  Raw data files filepath
Raw_File_Folder = easygui.diropenbox("Select The Directory Of Raw Data Files")
Summary_File_Folder = easygui.diropenbox("Select The Directory To Save The Summary File To")

RLU_Filename_Assay_df = pd.DataFrame()

RLU_Filename_dict = {}
RLU_Filename_Time_Start_End_dict = {}
RLU_Concurrent_Filenames_dict = {}
RLU_Concurrent_Assays_dict = {}
Sample_Cartridge_dict = {}

#  Create master curve level string list
char_list = ['s', 'sd', 'ss', 'st', 'std']
MC_Level_ID_list = []

for char in char_list:
    for i in range(0, 7):
        MC_Level_ID_list.append(char + str(i))

RLU_Filename_Time_Start_End_line_counter = 0
current_raw_file_line_counter = 0
#filecounter = 0
for path, subdirs, files in os.walk(Raw_File_Folder):

    #  Drop filenames that contain RLU (processed after summary files)
    files = [file for file in files if "RLU" not in file]

    for file in files:
        #  Check if file is NOT RLU file, save cartridge info
        #if file[0:3] != 'RLU':
        splitfile = file.split("_")
        Sample_Start_Time_list = splitfile[-7:-1]
        Sample_Start_Time = '_'.join(Sample_Start_Time_list)

        SampleID_found = False
        Cartridge_Lot_counter = 0
        Cartridge_Lot_ID = ''

        #  Read each line of file
        for line in open(Raw_File_Folder + '\\' + file, encoding='utf-16'):

            #  Remove new line character at the end of the current line
            line = line.rstrip("\n")

            splitline_tab = line.split("\t")

            if "SampleID" in line and not SampleID_found:
                splitline_colon = line.split(":", 1)
                SampleID = splitline_colon[1].strip()
                SampleID_found = True

            if len(splitline_tab) < 2:
                continue

            if splitline_tab[1] == "Type: Cartridge":
                splitline_colon = splitline_tab[2].split(":")
                Cartridge_Lot_SN = splitline_colon[1].strip()
                Cartridge_Lot_SN_split = Cartridge_Lot_SN.split("/")
                Cartridge_Lot = Cartridge_Lot_SN_split[0]
                Cartridge_SN = Cartridge_Lot_SN_split[1]
                Cartridge_Lot_counter += 1

        Sample_Cartridge_dict[SampleID + "___" + Sample_Start_Time] = [Cartridge_Lot, Cartridge_SN]


for path, subdirs, files in os.walk(Raw_File_Folder):
    for file in files:

        #  Drop filenames that don't contain RLU (already processed in previous step)
        files = [file for file in files if "RLU" in file]

        #  Check if file is RLU file, process
        if "RLU" in file and ".txt" in file:

            file_clean = file.strip()

            splitfile = file.split("_")
            Sample_Start_Time_list = splitfile[-7:-1]
            Sample_Start_Time = '_'.join(Sample_Start_Time_list)

            #  Initiate counter
            current_raw_file_line_counter = 0

            #  Read each line of file
            for line in open(Raw_File_Folder + '\\' + file, encoding='utf-8'):

                #  Remove new line character at the end of the current line
                line = line.rstrip("\n")

                #  Split line at each colon character
                splitline = line.split(":", 1)

                if splitline[0] == 'Date-Of-Report':
                    Test_End_Time_str = splitline[1].strip()

                elif splitline[0] == 'FileName':
                    Filename = splitline[1].strip()

                    #  Check if filename has extra characters at end (expect 'M' for AM/PM to be at certain position)
                    if Filename[-5] != 'M':

                        #  Drop extra characters
                        Filename = Filename[0: + Filename.rindex('_')]

                        #  Add leading zeros to hours if needed
                        Filename_split = Filename.split('_')
                        if len(Filename_split[-4]) == 1:
                            Filename_split[-4] = '0' + Filename_split[-4]
                            Filename = '_'.join(Filename_split)

                            Test_Start_Time_str = Filename[-22:]

                            #  Replace '_' character with appropriate characters for time
                            Test_Start_Time_str = Test_Start_Time_str.replace('_', "/", 2)
                            Test_Start_Time_str = Test_Start_Time_str.replace('_', ' ', 1)
                            Test_Start_Time_str = Test_Start_Time_str.replace('_', ':', 2)
                            Test_Start_Time_str = Test_Start_Time_str.replace('_', ' ')

                    else:
                        Filename_split = Filename.split('_')

                        #  Check if hours need leading zero added
                        if len(Filename_split[-4]) == 1:
                            Filename_split[-4] = '0' + Filename_split[-4]
                            Filename = '_'.join(Filename_split)

                        Test_Start_Time_str = Filename[-26:-4]
                        Test_Start_Time_str = Test_Start_Time_str.split('M')
                        Test_Start_Time_str = Test_Start_Time_str[0] + 'M'

                        #  Replace '_' character with appropriate characters for time
                        Test_Start_Time_str = Test_Start_Time_str.replace('_',"/", 2)
                        Test_Start_Time_str = Test_Start_Time_str.replace('_', ' ', 1)
                        Test_Start_Time_str = Test_Start_Time_str.replace('_', ':', 2)
                        Test_Start_Time_str = Test_Start_Time_str.replace('_', ' ')

                elif splitline[0] == 'Test':
                    Test = splitline[1].split('Test')
                    Test = Test[0].strip()

                elif splitline[0] == 'SampleID':
                    SampleID = splitline[1].strip()

                elif splitline[0] == 'RLUAverage':
                    RLUAverage = splitline[1].strip()

                    #  RLU average is the last value needed, break looping through lines once it's found
                    break

            #  Check if there was a full result file for current sample, save cartridge info if true
            if SampleID + '___' + Sample_Start_Time in Sample_Cartridge_dict:
                Reagent_Cartridge_Lot = Sample_Cartridge_dict[SampleID + '___' + Sample_Start_Time][0]
                Reagent_Cartridge_SN = Sample_Cartridge_dict[SampleID + '___' + Sample_Start_Time][1]
            else:
                Reagent_Cartridge_Lot = ''
                Reagent_Cartridge_SN = ''

            RLU_Filename_dict[file_clean] = [SampleID] + [Test] + [RLUAverage] + [Reagent_Cartridge_Lot] + [Reagent_Cartridge_SN] + [Test_Start_Time_str] + [Test_End_Time_str]

            RLU_Filename_Time_Start_End_dict[RLU_Filename_Time_Start_End_line_counter] = [file_clean] + [SampleID] + [Test] + [RLUAverage] + ['Start'] + [Test_Start_Time_str]
            RLU_Filename_Time_Start_End_line_counter += 1

            RLU_Filename_Time_Start_End_dict[RLU_Filename_Time_Start_End_line_counter] = [file_clean] + [SampleID] + [Test] + [RLUAverage] + ['End'] + [Test_End_Time_str]
            RLU_Filename_Time_Start_End_line_counter += 1
            #filecounter += 1

            #if filecounter == 100:
            #    break

#  Create dataframes from dictionaries
RLU_Filename_df = pd.DataFrame.from_dict(RLU_Filename_dict, "index", columns=['SampleID', 'Test', 'RLUAverage', 'Reagent_Cartridge_Lot', 'Reagent_Cartridge_SN', 'Test_Start_Time', 'Test_End_Time'])
RLU_Filename_Time_Start_End_df = pd.DataFrame.from_dict(RLU_Filename_Time_Start_End_dict, "index", columns=['Filename', 'SampleID', 'Test', 'RLUAverage', 'Start-End', 'Time'])

#  Convert RLU columns to numeric
RLU_Filename_df['RLUAverage'] = pd.to_numeric(RLU_Filename_df['RLUAverage'], errors='coerce')
RLU_Filename_Time_Start_End_df['RLUAverage'] = pd.to_numeric(RLU_Filename_Time_Start_End_df['RLUAverage'], errors='coerce')

#  Convert time columns to datetime
RLU_Filename_Time_Start_End_df['Time'] = pd.to_datetime(RLU_Filename_Time_Start_End_df['Time'], errors='coerce')
RLU_Filename_df['Test_Start_Time'] = pd.to_datetime(RLU_Filename_df['Test_Start_Time'], errors='coerce')
RLU_Filename_df['Test_End_Time'] = pd.to_datetime(RLU_Filename_df['Test_End_Time'], errors='coerce')

#  Sort values in start-end df by time column
RLU_Filename_Time_Start_End_df = RLU_Filename_Time_Start_End_df.sort_values(by=['Time'])

#  Get list of unique filenames in start-end df
Filename_Unique_list = RLU_Filename_Time_Start_End_df['Filename'].unique()

#  Get list of all filenames in start-end df
Filename_Full_list = RLU_Filename_Time_Start_End_df['Filename'].to_list()

#  Loop through each filename to get other filenames that were active at the same time
#  Skip filenames that don't contain expected master curve level identifiers
for Filename in Filename_Unique_list:

    Concurrent_Assay_list = []

    #  Get start and end rows of current filename
    Filename_indices = [indice for indice, file in enumerate(Filename_Full_list) if file == Filename]

    #  Get filenames of all filenames between current file indices
    Concurrent_Filenames_list = Filename_Full_list[Filename_indices[0] + 1: Filename_indices[1]]

    #  Remove duplicates (list contains start and end rows for each filename)
    Concurrent_Filenames_set = set(Concurrent_Filenames_list)
    Concurrent_Filenames_Unique_list = sorted(list(Concurrent_Filenames_set))

    #  Get assays run between current file indices
    for filename in Concurrent_Filenames_Unique_list:
        Concurrent_Assay_list.append(RLU_Filename_dict[filename][1])

    #  Create list of unique assays from full list
    Concurrent_Assay_set = set(Concurrent_Assay_list)
    Concurrent_Assay_Unique_list = sorted(list(Concurrent_Assay_set))

    #  Save filename start time and all concurrent assays to dictionary
    if Filename in RLU_Concurrent_Assays_dict:
        RLU_Concurrent_Assays_dict[Filename] = RLU_Concurrent_Assays_dict[Filename] + Concurrent_Assay_Unique_list
    else:
        RLU_Concurrent_Assays_dict[Filename] = [RLU_Filename_dict[Filename][5]] + Concurrent_Assay_Unique_list

    #  Save filename start time and all concurrent filenames to dictionary
    if Filename in RLU_Concurrent_Filenames_dict:
        RLU_Concurrent_Filenames_dict[Filename] = RLU_Concurrent_Filenames_dict[Filename] + Concurrent_Filenames_Unique_list
    else:
        RLU_Concurrent_Filenames_dict[Filename] = [RLU_Filename_dict[Filename][5]] + Concurrent_Filenames_Unique_list

    #  Get filenames and assays of all filenames that completely contain current filename
    #  (filename started before current and ended after current)
    #  Perform by saving current filename to all filenames found to be concurrent (delete duplicates later)
    for Filename2 in Concurrent_Filenames_Unique_list:

        #  Skip if filenames are the same
        if Filename2 == Filename:
            continue

        if Filename2 in RLU_Concurrent_Filenames_dict:
            RLU_Concurrent_Filenames_dict[Filename2] = RLU_Concurrent_Filenames_dict[Filename2] + [Filename]

        else:
            RLU_Concurrent_Filenames_dict[Filename2] = [RLU_Filename_dict[Filename2][5]] + [Filename]

        if Filename2 in RLU_Concurrent_Assays_dict:
            RLU_Concurrent_Assays_dict[Filename2] = RLU_Concurrent_Assays_dict[Filename2] + [RLU_Filename_dict[Filename][1]]

        else:
            RLU_Concurrent_Assays_dict[Filename2] = [RLU_Filename_dict[Filename2][5]] + [RLU_Filename_dict[Filename][1]]

#  Remove duplicate filenames and assays from each dictionary
for filename in RLU_Concurrent_Assays_dict:
    Start_Time = RLU_Concurrent_Assays_dict[filename][0]
    Assay_set = set(RLU_Concurrent_Assays_dict[filename][1:])
    RLU_Concurrent_Assays_dict[filename] = [Start_Time] + sorted(list(Assay_set))

for filename in RLU_Concurrent_Filenames_dict:
    Start_Time = RLU_Concurrent_Filenames_dict[filename][0]
    Filename_set = set(RLU_Concurrent_Filenames_dict[filename][1:])
    RLU_Concurrent_Filenames_dict[filename] = [Start_Time] + sorted(list(Filename_set))


RLU_Concurrent_Filenames_df = pd.DataFrame.from_dict(RLU_Concurrent_Filenames_dict, "index")
RLU_Concurrent_Filenames_df.rename(columns={0: 'Filename_Start_Time'})

RLU_Concurrent_Assay_df = pd.DataFrame.from_dict(RLU_Concurrent_Assays_dict, "index")
RLU_Concurrent_Assay_df.rename(columns={0: 'Assay_Start_Time'})

#RLU_Concurrent_Filenames_df.to_csv(Summary_File_Folder + '//Concurrent_Filenames_' + timestr + '.txt', index=True)
#RLU_Concurrent_Assay_df.to_csv(Summary_File_Folder + '//Concurrent_Assays_' + timestr + '.txt', index=True)
#RLU_Filename_Time_Start_End_df.to_csv(Summary_File_Folder + '//Summary_Start-End_' + timestr + '.txt', index=False)
#RLU_Filename_df.to_csv(Summary_File_Folder + '//Summary_' + timestr + '.txt', index=False)

#  Sort values in summary dataframe by test start time
RLU_Filename_df = RLU_Filename_df.sort_values(by=['Test_Start_Time'])

#  Create excel file
with pd.ExcelWriter(Summary_File_Folder + "\\" +
                    'RLU-File-Summary_' +
                    timestr + ".xlsx", engine='xlsxwriter') as writer:
    RLU_Filename_df.to_excel(writer, sheet_name='Summary', index=True, header=True)
    RLU_Concurrent_Assay_df.to_excel(writer, sheet_name='Concurrent_Assays', index=True, header=True)
    RLU_Concurrent_Filenames_df.to_excel(writer, sheet_name='Concurrent_Files', index=True, header=True)



