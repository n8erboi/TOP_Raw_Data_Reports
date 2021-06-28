import time
import easygui
import os
import pandas as pd
import numpy as np
import io
from datetime import datetime

#  This script compiles ACL TOP raw data files into a single excel file

#  Create timestamp for output filename
timestr = time.strftime("%Y_%m_%d-%H%M%S")

#  Raw data files filepath
Raw_File_Folder = easygui.diropenbox("Select The Directory Of Raw Data Files")
Summary_File_Folder = easygui.diropenbox("Select The Directory To Save The Summary File To")

Raw_File_Temporary_dict = {}
Raw_File_Overall_dict = {}
Raw_File_Overall_Alarms_dict = {}

AM_Data_Unstacked_df = pd.DataFrame()
Raw_Unstacked_df = pd.DataFrame()
Normalized_Unstacked_df = pd.DataFrame()
Deriv1_Unstacked_df = pd.DataFrame()
Deriv2_Unstacked_df = pd.DataFrame()

Column_Labels_list = ['SampleID',
                      'Test_Code',
                      'Test_Date',
                      'Time',
                      'AM_Data',
                      'Flags',
                      'Raw',
                      'Normalized',
                      'Deriv1',
                      'Deriv2',
                      'Alarm_1',
                      'Alarm_2',
                      'Alarm_3',
                      'Alarm_4',
                      'Alarm_5',
                      'Alarm_6',
                      'Alarm_7',
                      'Alarm_8',
                      'Alarm_9',
                      'Alarm_10',
                      'Alarm_11',
                      'Alarm_12',
                      'Alarm_13',
                      'Alarm_14',
                      'Alarm_15',
                      'Alarm_16',
                      'Alarm_17',
                      'Alarm_18',
                      'Alarm_19',
                      'Alarm_20']

Overall_Error_Code_list = []

current_raw_file_line_counter = 0
overall_raw_file_line_counter = 0
for path, subdirs, files in os.walk(Raw_File_Folder):
    for file in files:

        if ".txt" in file:

            #  Initiate counter
            current_raw_file_line_counter = 0

            #  Initiate line reader variable
            Save_Lines = False

            SampleID = ''
            Test_Code = ''
            Test_Date = ''
            Test_Time = ''

            SampleID_Found = False
            Test_Code_Found = False
            Test_Date_Found = False
            Start_Found = False
            Save_Errors = False

            Current_Error_Code_list = []

            FileType = 'utf-8'
            #  Check if file is utf-16-le
            try:
                for line in open(Raw_File_Folder + '\\' + file, encoding=FileType):
                    continue
            except UnicodeDecodeError:
                FileType = 'utf-16'


            #  Read each line of file
            if FileType == 'utf-8':
                for line in open(Raw_File_Folder + '\\' + file, encoding=FileType):

                    #  Remove new line character at the end of the current line
                    line = line.rstrip("\n")

                    #  Split line at each tab character, remove last element empty
                    splitline = line.split("\t")
                    if splitline[-1] == '' and len(splitline) != 1:
                        del splitline[-1]

                    #  Skip lines one element long, unless it's equal to #End
                    if len(splitline) == 1 and splitline[0] != "#End":
                        continue

                    #  Find test info
                    if splitline[0] == "#SampleID:" and not SampleID_Found:
                        SampleID = splitline[1]
                        SampleID_Found = True

                    if len(splitline) > 3:
                        if splitline[2] == "#Test:" and not Test_Code_Found:
                            Test_Code = splitline[3]
                            Test_Code_Found = True

                        elif splitline[2] == "Order Date:" and not Test_Date_Found:
                            Test_Date = splitline[3]
                            Test_Date_Found = True

                    if splitline[0] == "#Time":
                        Save_Lines = True

                    if splitline[0] == "#End":
                        Save_Lines = False

                    if splitline[0] == "-----":
                        continue

                    if line == 'Errors and Warnings:':
                        Save_Errors = True

                    if Save_Errors and len(splitline[1]) == 2:
                        Current_Error_Code_list.append(splitline[1] + '_' + splitline[2])
                        Overall_Error_Code_list.append(splitline[1] + '_' + splitline[2])

                    if Save_Lines and "#" not in splitline[0]:

                        if len(splitline[0]) == 1:
                            splitline[0] = '000' + splitline[0]
                        elif len(splitline[0]) == 2:
                            splitline[0] = '00' + splitline[0]
                        elif len(splitline[0]) == 3:
                            splitline[0] = '0' + splitline[0]

                        if 'D' in splitline[0]:
                            splitline[0] = '0' + splitline[0]

                        #  Add placeholder for Deriv2 if line is missing it
                        if len(splitline) == 6:
                            Raw_File_Temporary_dict[current_raw_file_line_counter] = splitline + ['']
                        else:
                            Raw_File_Temporary_dict[current_raw_file_line_counter] = splitline

                        current_raw_file_line_counter += 1

                #  Update overall dictionary with current dictionary values plus test info
                for key in Raw_File_Temporary_dict:
                    Raw_File_Overall_dict[overall_raw_file_line_counter] = [SampleID] + [Test_Code] + [Test_Date] + Raw_File_Temporary_dict[key]
                    overall_raw_file_line_counter += 1

                #  Update error tracking dictionary
                Sample_Test_Date = SampleID + '_' + Test_Code + '_' + Test_Date
                Raw_File_Overall_Alarms_dict[Sample_Test_Date] = Current_Error_Code_list

            else:
                for line in open(Raw_File_Folder + '\\' + file, encoding=FileType):
                    Split_Tabs = False

                    if "\t" in line:
                        Split_Tabs = True

                    splitline = line.split("\t")
                    line = ' '.join(splitline)

                    #  Remove new line character at the end of the current line
                    line = line.rstrip("\n")

                    #  Split line at each tab character, remove last element empty
                    splitline = line.split(" ")
                    if splitline[-1] == '' and len(splitline) != 1:
                        del splitline[-1]

                    #  Skip lines one element long, unless it's equal to #End
                    if len(splitline) == 1 and splitline[0] != "#End":
                        continue

                    #  Find test info
                    if splitline[0] == "#SampleID:" and not SampleID_Found:
                        SampleID = ' '.join(splitline[1:])
                        SampleID_Found = True

                    if len(splitline) > 3:
                        if splitline[0] == "#Test:" and not Test_Code_Found:
                            Test_Code = ' '.join(splitline[1:-3])
                            Test_Code_Found = True

                        elif splitline[3] == "Order" and splitline[4] == 'Date:' and not Test_Date_Found:
                            Test_Date = splitline[5]
                            Test_Time = splitline[6]
                            Test_Date = Test_Date + ' ' + Test_Time
                            Test_Date_Found = True

                    if splitline[0] == "#Time" and not Start_Found:
                        Save_Lines = True
                        Start_Found = True

                    if splitline[0] == "#End":
                        Save_Lines = False

                    if splitline[0] == "-----":
                        continue

                    if line == 'Errors and Warnings:':
                        Save_Errors = True

                    if Save_Errors and len(splitline[1]) == 2:
                        Current_Error_Code_list.append(splitline[1] + '_' + splitline[2])
                        Overall_Error_Code_list.append(splitline[1] + '_' + splitline[2])

                    if Save_Lines and "#" not in splitline[0]:

                        #  Add leading zeros if needed
                        if len(splitline[0]) == 1:
                            splitline[0] = '000' + splitline[0]
                        elif len(splitline[0]) == 2:
                            splitline[0] = '00' + splitline[0]
                        elif len(splitline[0]) == 3:
                            splitline[0] = '0' + splitline[0]

                        if 'D' in splitline[0]:
                            splitline[0] = '0' + splitline[0]

                        Raw_File_Temporary_dict[current_raw_file_line_counter] = splitline

                        current_raw_file_line_counter += 1

                #  Update overall dictionary with current dictionary values plus test info
                for key in Raw_File_Temporary_dict:
                    Raw_File_Overall_dict[overall_raw_file_line_counter] = [SampleID] + [Test_Code] + [Test_Date] + \
                                                                           Raw_File_Temporary_dict[key]
                    overall_raw_file_line_counter += 1

                #  Update error tracking dictionary
                Sample_Test_Date = SampleID + '_' + Test_Code + '_' + Test_Date
                Raw_File_Overall_Alarms_dict[Sample_Test_Date] = Current_Error_Code_list

#  Create dataframe from overall dictionary
#if FileType == 'utf-8':
Raw_File_Overall_df = pd.DataFrame.from_dict(Raw_File_Overall_dict, "index")

#  Rename dataframe columns (number of parameter columns will vary, initial columns always the same)
RawData_Header_list = []
for column in Raw_File_Overall_df:
    RawData_Header_list.append(Column_Labels_list[column])

Raw_File_Overall_df.columns = RawData_Header_list

if 'Deriv1' not in Raw_File_Overall_df.columns:
    Raw_File_Overall_df['Deriv1'] = ''

if 'Deriv2' not in Raw_File_Overall_df.columns:
    Raw_File_Overall_df['Deriv2'] = ''

#  Convert certain columns to numeric
Raw_File_Overall_df['AM_Data'] = pd.to_numeric(Raw_File_Overall_df['AM_Data'], errors='coerce')
Raw_File_Overall_df['Raw'] = pd.to_numeric(Raw_File_Overall_df['Raw'], errors='coerce')
Raw_File_Overall_df['Normalized'] = pd.to_numeric(Raw_File_Overall_df['Normalized'], errors='coerce')
Raw_File_Overall_df['Deriv1'] = pd.to_numeric(Raw_File_Overall_df['Deriv1'], errors='coerce')
Raw_File_Overall_df['Deriv2'] = pd.to_numeric(Raw_File_Overall_df['Deriv2'], errors='coerce')

#  Create unstacked dataframes for each value
#  Create separate dataframes
AM_Data_Unstacked_df['ID_Column'] = Raw_File_Overall_df['SampleID'] + '_' + Raw_File_Overall_df['Test_Code'] + '_' + Raw_File_Overall_df['Test_Date']
AM_Data_Unstacked_df['Data'] = Raw_File_Overall_df['AM_Data']
AM_Data_Unstacked_df['Time'] = Raw_File_Overall_df['Time']
AM_Data_Unstacked_df['Time'] = AM_Data_Unstacked_df['Time'].str.replace('D', '')
AM_Data_Unstacked_df['Time'] = pd.to_numeric(AM_Data_Unstacked_df['Time'], errors='coerce')

Raw_Unstacked_df['ID_Column'] = Raw_File_Overall_df['SampleID'] + '_' + Raw_File_Overall_df['Test_Code'] + '_' + Raw_File_Overall_df['Test_Date']
Raw_Unstacked_df['Data'] = Raw_File_Overall_df['Raw']
Raw_Unstacked_df['Time'] = Raw_File_Overall_df['Time']
Raw_Unstacked_df['Time'] = Raw_Unstacked_df['Time'].str.replace('D', '')
Raw_Unstacked_df['Time'] = pd.to_numeric(Raw_Unstacked_df['Time'], errors='coerce')

Normalized_Unstacked_df['ID_Column'] = Raw_File_Overall_df['SampleID'] + '_' + Raw_File_Overall_df['Test_Code'] + '_' + Raw_File_Overall_df['Test_Date']
Normalized_Unstacked_df['Data'] = Raw_File_Overall_df['Normalized']
Normalized_Unstacked_df['Time'] = Raw_File_Overall_df['Time']
Normalized_Unstacked_df['Time'] = Normalized_Unstacked_df['Time'].str.replace('D', '')
Normalized_Unstacked_df['Time'] = pd.to_numeric(Normalized_Unstacked_df['Time'], errors='coerce')

Deriv1_Unstacked_df['ID_Column'] = Raw_File_Overall_df['SampleID'] + '_' + Raw_File_Overall_df['Test_Code'] + '_' + Raw_File_Overall_df['Test_Date']
Deriv1_Unstacked_df['Data'] = Raw_File_Overall_df['Deriv1']
Deriv1_Unstacked_df['Time'] = Raw_File_Overall_df['Time']
Deriv1_Unstacked_df['Time'] = Deriv1_Unstacked_df['Time'].str.replace('D', '')
Deriv1_Unstacked_df['Time'] = pd.to_numeric(Deriv1_Unstacked_df['Time'], errors='coerce')

Deriv2_Unstacked_df['ID_Column'] = Raw_File_Overall_df['SampleID'] + '_' + Raw_File_Overall_df['Test_Code'] + '_' + Raw_File_Overall_df['Test_Date']
Deriv2_Unstacked_df['Data'] = Raw_File_Overall_df['Deriv2']
Deriv2_Unstacked_df['Time'] = Raw_File_Overall_df['Time']
Deriv2_Unstacked_df['Time'] = Deriv2_Unstacked_df['Time'].str.replace('D', '')
Deriv2_Unstacked_df['Time'] = pd.to_numeric(Deriv2_Unstacked_df['Time'], errors='coerce')

#  Pivot and sort each dataframe
AM_Data_Unstacked_df = AM_Data_Unstacked_df.pivot(index='Time', columns='ID_Column', values='Data')
AM_Data_Unstacked_df.sort_index()

Raw_Unstacked_df = Raw_Unstacked_df.pivot(index='Time', columns='ID_Column', values='Data')
Raw_Unstacked_df.sort_index()

Normalized_Unstacked_df = Normalized_Unstacked_df.pivot(index='Time', columns='ID_Column', values='Data')
Normalized_Unstacked_df.sort_index()

Deriv1_Unstacked_df = Deriv1_Unstacked_df.pivot(index='Time', columns='ID_Column', values='Data')
Deriv1_Unstacked_df.sort_index()

Deriv2_Unstacked_df = Deriv2_Unstacked_df.pivot(index='Time', columns='ID_Column', values='Data')
Deriv2_Unstacked_df.sort_index()

#  Count alarms and create dataframe
Overall_Unique_Error_Code_list = sorted(list(set(Overall_Error_Code_list)))

Alarm_Tracking_dict = {}
for sample_key in Raw_File_Overall_Alarms_dict:
    Alarm_Tracking_dict[sample_key] = ''
    for alarm in Overall_Unique_Error_Code_list:
        if Overall_Unique_Error_Code_list.index(alarm) == 0:
            if alarm in Raw_File_Overall_Alarms_dict[sample_key]:
                Alarm_Tracking_dict[sample_key] = [True]
            else:
                Alarm_Tracking_dict[sample_key] = [False]
        else:
            if alarm in Raw_File_Overall_Alarms_dict[sample_key]:
                Alarm_Tracking_dict[sample_key] = Alarm_Tracking_dict[sample_key] + [True]
            else:
                Alarm_Tracking_dict[sample_key] = Alarm_Tracking_dict[sample_key] + [False]

Overall_Alarm_df = pd.DataFrame.from_dict(Alarm_Tracking_dict, "index")

Overall_Alarm_df.columns = Overall_Unique_Error_Code_list

#  Create excel file
with pd.ExcelWriter(Summary_File_Folder + "\\" +
                    'Raw-Data-File-Summary_' +
                    timestr + ".xlsx", engine='xlsxwriter') as writer:
    Raw_File_Overall_df.to_excel(writer, sheet_name='All_Data', index=False, header=True)
    Overall_Alarm_df.to_excel(writer, sheet_name='Alarms', index=True, header=True)
    AM_Data_Unstacked_df.to_excel(writer, sheet_name='AM_Data', index=True, header=True)
    Raw_Unstacked_df.to_excel(writer, sheet_name='Raw_Data', index=True, header=True)
    Normalized_Unstacked_df.to_excel(writer, sheet_name='Normalized_Data', index=True, header=True)
    Deriv1_Unstacked_df.to_excel(writer, sheet_name='Deriv1_Data', index=True, header=True)
    Deriv2_Unstacked_df.to_excel(writer, sheet_name='Deriv2_Data', index=True, header=True)



