import time
import easygui
import os
import pandas as pd
import math
from scipy import stats
import matplotlib.pyplot as plt
import numpy as np
import io
from datetime import date
from datetime import datetime
import datetime

#  This script compiles ACL TOP raw RLU files into a single excel file and calculates metrics based on current
#  and user-entered values

#  Create timestamp for output filename
timestr = time.strftime("%Y_%m_%d-%H%M%S")

#  Define function to convert dates to excel format
def DateTime_to_ExcelDate(Original_Date):
    Excel_Start_Date = datetime.datetime(1899, 12, 30)
    Date_Delta = Original_Date - Excel_Start_Date
    Updated_Date = Date_Delta.dt.days + Date_Delta.dt.seconds / 86400
    return Updated_Date

#  Create function to convert column numbers to letters for excel
def column_number_to_string(n):
    #  Create string variable
    string = ""
    n += 1
    while n > 0:
        #  Divide column number by 26, save result as new n and remainder
        n, remainder = divmod(n - 1, 26)

        #  Convert remainder number to string, append to existing string
        string = chr(65 + remainder) + string
    return string

#  Raw data files filepath
Raw_File_Folder = easygui.diropenbox("Select The Directory Of Raw Data Files")
Summary_File_Folder = easygui.diropenbox("Select The Directory To Save The Summary File To")

#  Get user-defined time ranges and data reduction limits

TimeRanges_MEB_msg = "Enter the investigational set of time ranges or leave blank"
TimeRanges_MEB_title = "Investigational Time Ranges"
TimeRanges_MEB_FieldNames = ['Background 1 Check - Start (Current: 4)',
                             'Background 1 Check - End (Current: 85)',
                             'Background 2 Check - Start (Current: 101)',
                             'Background 2 Check - End (Current: 182)',
                             'Background 3 Check - Start (Current: 86)',
                             'Background 3 Check - End (Current: 100)',
                             'Slope Check - Start (Current: 216)',
                             'Slope Check - End (Current: 254)',
                             'Spike Check #1 - Start (Current: 4)',
                             'Spike Check #1 - End (Current: 206)',
                             'Spike Check #2 - Start (Current: 207)',
                             'Spike Check #2 - End (Current: 302)']

TimeRanges_MEB_FieldValues = easygui.multenterbox(TimeRanges_MEB_msg, TimeRanges_MEB_title, TimeRanges_MEB_FieldNames)

if all(value == '' for value in TimeRanges_MEB_FieldValues):
    User_TimeRanges = False
    User_TimeRanges_Background1_Enabled = False
    User_TimeRanges_Background2_Enabled = False
    User_TimeRanges_Background3_Enabled = False
    User_TimeRanges_Slope_Enabled = False
    User_TimeRanges_Spike1_Enabled = False
    User_TimeRanges_Spike2_Enabled = False

else:
    User_TimeRanges = True

    if TimeRanges_MEB_FieldValues[0] != '' and TimeRanges_MEB_FieldValues[1] != '':
        User_TimeRanges_Background1_Enabled = True
    else:
        User_TimeRanges_Background1_Enabled = False

    if TimeRanges_MEB_FieldValues[2] != '' and TimeRanges_MEB_FieldValues[3] != '':
        User_TimeRanges_Background2_Enabled = True
    else:
        User_TimeRanges_Background2_Enabled = False

    if TimeRanges_MEB_FieldValues[4] != '' and TimeRanges_MEB_FieldValues[5] != '':
        User_TimeRanges_Background3_Enabled = True
    else:
        User_TimeRanges_Background3_Enabled = False

    if TimeRanges_MEB_FieldValues[6] != '' and TimeRanges_MEB_FieldValues[7] != '':
        User_TimeRanges_Slope_Enabled = True
    else:
        User_TimeRanges_Slope_Enabled = False

    if TimeRanges_MEB_FieldValues[8] != '' and TimeRanges_MEB_FieldValues[9] != '':
        User_TimeRanges_Spike1_Enabled = True
    else:
        User_TimeRanges_Spike1_Enabled = False

    if TimeRanges_MEB_FieldValues[10] != '' and TimeRanges_MEB_FieldValues[11] != '':
        User_TimeRanges_Spike2_Enabled = True
    else:
        User_TimeRanges_Spike2_Enabled = False


DRLimits_MEB_msg = "Enter the investigational set of DR limits for background readings or leave blank"
DRLimits_MEB_title = "Data Reduction Investigational Limits"
DRLimits_MEB_FieldNames = ['Background 1 Upper Limit (Current: 150)',
                           'Background 2 Upper Limit (Current: 150)',
                           'Background 3 Upper Limit (Current: 150)',
                           'Slope Lower Limit (Current: -1.1)',
                           'Slope Upper Limit (Current: -0.4)']

DRLimits_MEB_FieldValues = easygui.multenterbox(DRLimits_MEB_msg, DRLimits_MEB_title, DRLimits_MEB_FieldNames)

if all(value == '' for value in DRLimits_MEB_FieldValues):
    User_DRLimits = False
    User_DRLimits_Background1_Enabled = False
    User_DRLimits_Background2_Enabled = False
    User_DRLimits_Background3_Enabled = False
    User_DRLimits_Slope_Enabled = False

else:
    User_DRLimits = True

    if DRLimits_MEB_FieldValues[0] != '':
        User_DRLimits_Background1_Enabled = True
    else:
        User_DRLimits_Background1_Enabled = False

    if DRLimits_MEB_FieldValues[1] != '':
        User_DRLimits_Background2_Enabled = True
    else:
        User_DRLimits_Background2_Enabled = False

    if DRLimits_MEB_FieldValues[2] != '':
        User_DRLimits_Background3_Enabled = True
    else:
        User_DRLimits_Background3_Enabled = False

    if DRLimits_MEB_FieldValues[3] != '' and DRLimits_MEB_FieldValues[4] != '':
        User_DRLimits_Slope_Enabled = True
    else:
        User_DRLimits_Slope_Enabled = False

#  Save current and time ranges and DR values
Current_Background1_Start = 4
Current_Background1_End = 85
Current_Background2_Start = 101
Current_Background2_End = 182
Current_Background3_Start = 86
Current_Background3_End = 100
Current_Slope_Start = 216
Current_Slope_End = 254
Current_Spike1_Start = 4
Current_Spike1_End = 206
Current_Spike2_Start = 207
Current_Spike2_End = 302

Current_Background1_UpperLimit = 150
Current_Background2_UpperLimit = 150
Current_Background3_UpperLimit = 150
Current_Slope_LowerLimit = -1.10
Current_Slope_UpperLimit = -0.40

#  Create lists for current time ranges (used to find index rows in each file)
Current_Background1_Index_list = [*range(Current_Background1_Start, Current_Background1_End + 1)]
Current_Background2_Index_list = [*range(Current_Background2_Start, Current_Background2_End + 1)]
Current_Background3_Index_list = [*range(Current_Background3_Start, Current_Background3_End + 1)]
Current_Slope_Index_list = [*range(Current_Slope_Start, Current_Slope_End + 1)]
Current_Slope_ms_list = [i * 25 for i in Current_Slope_Index_list]
Current_Spike1_Index_list = [*range(Current_Spike1_Start, Current_Spike1_End + 1)]
Current_Spike2_Index_list = [*range(Current_Spike2_Start, Current_Spike2_End + 1)]

#  Repeat if user entered time ranges
if User_TimeRanges:

    if User_TimeRanges_Background1_Enabled:
        User_Background1_Start = float(TimeRanges_MEB_FieldValues[0])
        User_Background1_End = float(TimeRanges_MEB_FieldValues[1])
        User_Background1_Index_list = [*range(int(User_Background1_Start), int(User_Background1_End) + 1)]
    else:
        User_Background1_Index_list = []

    if User_TimeRanges_Background2_Enabled:
        User_Background2_Start = float(TimeRanges_MEB_FieldValues[2])
        User_Background2_End = float(TimeRanges_MEB_FieldValues[3])
        User_Background2_Index_list = [*range(int(User_Background2_Start), int(User_Background2_End) + 1)]
    else:
        User_Background2_Index_list = []

    if User_TimeRanges_Background3_Enabled:
        User_Background3_Start = float(TimeRanges_MEB_FieldValues[4])
        User_Background3_End = float(TimeRanges_MEB_FieldValues[5])
        User_Background3_Index_list = [*range(int(User_Background3_Start), int(User_Background3_End) + 1)]
    else:
        User_Background3_Index_list = []

    if User_TimeRanges_Slope_Enabled:
        User_Slope_Start = float(TimeRanges_MEB_FieldValues[6])
        User_Slope_End = float(TimeRanges_MEB_FieldValues[7])
        User_Slope_Index_list = [*range(int(User_Slope_Start), int(User_Slope_End) + 1)]
        User_Slope_ms_list = [i * 25 for i in User_Slope_Index_list]
    else:
        User_Slope_Index_list = []
        User_Slope_ms_list = []

    if User_TimeRanges_Spike1_Enabled:
        User_Spike1_Start = float(TimeRanges_MEB_FieldValues[8])
        User_Spike1_End = float(TimeRanges_MEB_FieldValues[9])
        User_Spike1_Index_list = [*range(int(User_Spike1_Start), int(User_Spike1_End) + 1)]
    else:
        User_Spike1_Index_list = []


    if User_TimeRanges_Spike2_Enabled:
        User_Spike2_Start = float(TimeRanges_MEB_FieldValues[10])
        User_Spike2_End = float(TimeRanges_MEB_FieldValues[11])
        User_Spike2_Index_list = [*range(int(User_Spike2_Start), int(User_Spike2_End) + 1)]
    else:
        User_Spike2_Index_list = []



if User_DRLimits:

    if User_DRLimits_Background1_Enabled:
        User_Background1_UpperLimit = float(DRLimits_MEB_FieldValues[0])

    if User_DRLimits_Background2_Enabled:
        User_Background2_UpperLimit = float(DRLimits_MEB_FieldValues[1])

    if User_DRLimits_Background3_Enabled:
        User_Background3_UpperLimit = float(DRLimits_MEB_FieldValues[2])

    if User_DRLimits_Slope_Enabled:
        User_Slope_LowerLimit = float(DRLimits_MEB_FieldValues[3])
        User_Slope_UpperLimit = float(DRLimits_MEB_FieldValues[4])

RLU_Filename_Assay_df = pd.DataFrame()

RLU_DRValues_User_dict = {}
RLU_DRValues_Current_dict = {}

RLU_Filename_Time_Start_End_line_counter = 0
current_raw_file_line_counter = 0

for path, subdirs, files in os.walk(Raw_File_Folder):
    for file in files:

        if "RLU" in file and ".txt" in file:
            print(file)
            #  Create lists to save RLU values for each parameter
            Current_Background1_RLU_list = []
            Current_Background2_RLU_list = []
            Current_Background3_RLU_list = []
            Current_Slope_RLU_list = []
            Current_Slope_lnRLU_list = []
            Current_Spike1_RLU_list = []
            Current_Spike2_RLU_list = []

            User_Background1_RLU_list = []
            User_Background2_RLU_list = []
            User_Background3_RLU_list = []
            User_Slope_RLU_list = []
            User_Slope_lnRLU_list = []
            User_Spike1_RLU_list = []
            User_Spike2_RLU_list = []

            file_clean = file.strip()

            #  Initiate counter
            current_raw_file_line_counter = 0

            #  Initiate variable to check when index values start in the file
            StartIndex = False

            #  Initiate variable to check if any RLU values are zero (to prevent error when calculating ln(RLU))
            Zero_RLU_Value_Found = False

            #  Read each line of file
            for line in open(Raw_File_Folder + '\\' + file, encoding='utf-8'):

                #  Remove new line character at the end of the current line
                line = line.rstrip("\n")

                #  Check if all data has been found in the file, break loop if true
                if StartIndex and line == '':
                    break

                #  Split line at each colon character
                splitline = line.split(":", 1)

                #  Split first element of splitline again at each tab character
                splitline_tab = splitline[0].split('\t')

                #  Check if current line is the first index value (0)
                if splitline_tab[0] == '0':
                    StartIndex = True

                #  Convert elements one and two to integers if the row is an index value
                if StartIndex:
                    splitline_tab[0] = int(splitline_tab[0])
                    splitline_tab[1] = int(splitline_tab[1])

                if splitline[0] == 'Date-Of-Report':
                    Date_split = splitline[1].split(' ')
                    Date = Date_split[1]

                elif splitline[0] == 'FileName':
                    Filename = splitline[1].strip()

                elif splitline[0] == 'Test':
                    Test = splitline[1].split('Test')
                    Test = Test[0].strip()

                elif splitline[0] == 'SampleID':
                    SampleID = splitline[1].strip()

                elif splitline[0] == 'RLUAverage':
                    RLUAverage = splitline[1].strip()

                #  Check if index value is in any of the user or current time ranges, save RLU values
                if splitline_tab[0] in Current_Background1_Index_list:
                    Current_Background1_RLU_list.append(splitline_tab[1])
                elif splitline_tab[0] in Current_Background2_Index_list:
                    Current_Background2_RLU_list.append(splitline_tab[1])
                elif splitline_tab[0] in Current_Background3_Index_list:
                    Current_Background3_RLU_list.append(splitline_tab[1])
                elif splitline_tab[0] in Current_Slope_Index_list:
                    Current_Slope_RLU_list.append(splitline_tab[1])
                elif splitline_tab[0] in Current_Spike1_Index_list:
                    Current_Spike1_RLU_list.append(splitline_tab[1])
                elif splitline_tab[0] in Current_Spike2_Index_list:
                    Current_Spike2_RLU_list.append(splitline_tab[1])

                if User_TimeRanges:
                    if splitline_tab[0] in User_Background1_Index_list:
                        User_Background1_RLU_list.append(splitline_tab[1])
                    elif splitline_tab[0] in User_Background2_Index_list:
                        User_Background2_RLU_list.append(splitline_tab[1])
                    elif splitline_tab[0] in User_Background3_Index_list:
                        User_Background3_RLU_list.append(splitline_tab[1])
                    elif splitline_tab[0] in User_Slope_Index_list:
                        User_Slope_RLU_list.append(splitline_tab[1])
                    elif splitline_tab[0] in User_Spike1_Index_list:
                        User_Spike1_RLU_list.append(splitline_tab[1])
                    elif splitline_tab[0] in User_Spike2_Index_list:
                        User_Spike2_RLU_list.append(splitline_tab[1])


            #  Calculate current and user metrics for current file
            CurrentFile_Current_Background1 = (40 * sum(Current_Background1_RLU_list)) / len(Current_Background1_RLU_list)
            CurrentFile_Current_Background2 = (40 * sum(Current_Background2_RLU_list)) / len(Current_Background2_RLU_list)
            CurrentFile_Current_Background3 = (40 * sum(Current_Background3_RLU_list)) / len(Current_Background3_RLU_list)

            #  Calculate natural log of raw RLU values within slope range
            for i in range(len(Current_Slope_RLU_list)):
                if Current_Slope_RLU_list[i] <= 0:
                    Zero_RLU_Value_Found = True
                    break
                else:
                    Current_Slope_lnRLU_list.append(math.log(Current_Slope_RLU_list[i]))

            #  Calculate slope
            if Zero_RLU_Value_Found:
                CurrentFile_Current_Slope = np.NaN
            else:
                CurrentFile_Current_Slope, Current_intercept, Current_r_value, Current_p_value, Current_std_error = stats.linregress(Current_Slope_ms_list, Current_Slope_lnRLU_list)
                CurrentFile_Current_Slope = 1000 * CurrentFile_Current_Slope
            #plt.plot(Current_Slope_ms_list, Current_Slope_lnRLU_list)
            #plt.show()

            CurrentFile_Current_Spike1 = (40 * sum(Current_Spike1_RLU_list)) / len(Current_Spike1_RLU_list)
            CurrentFile_Current_Spike2 = (40 * sum(Current_Spike2_RLU_list)) / len(Current_Spike2_RLU_list)

            if User_TimeRanges:
                if User_TimeRanges_Background1_Enabled:
                    CurrentFile_User_Background1 = (40 * sum(User_Background1_RLU_list)) / len(User_Background1_RLU_list)
                else:
                    CurrentFile_User_Background1 = 'N/A'

                if User_TimeRanges_Background2_Enabled:
                    CurrentFile_User_Background2 = (40 * sum(User_Background2_RLU_list)) / len(User_Background2_RLU_list)
                else:
                    CurrentFile_User_Background2 = 'N/A'

                if User_TimeRanges_Background3_Enabled:
                    CurrentFile_User_Background3 = (40 * sum(User_Background3_RLU_list)) / len(User_Background3_RLU_list)
                else:
                    CurrentFile_User_Background3 = 'N/A'

                #  Calculate natural log of raw RLU values within slope range
                if User_TimeRanges_Slope_Enabled:
                    for i in range(len(User_Slope_RLU_list)):
                        if User_Slope_RLU_list[i] <= 0:
                            Zero_RLU_Value_Found = True
                            break
                        else:
                            User_Slope_lnRLU_list.append(math.log(User_Slope_RLU_list[i]))

                    #  Calculate slope
                    if Zero_RLU_Value_Found:
                        CurrentFile_User_Slope = np.NaN
                    else:
                        CurrentFile_User_Slope, User_intercept, User_r_value, User_p_value, User_std_error = stats.linregress(User_Slope_ms_list, User_Slope_lnRLU_list)
                        CurrentFile_User_Slope = 1000 * CurrentFile_User_Slope
                else:
                    CurrentFile_User_Slope = 'N/A'

                if User_TimeRanges_Spike1_Enabled:
                    CurrentFile_User_Spike1 = (40 * sum(User_Spike1_RLU_list)) / len(User_Spike1_RLU_list)
                else:
                    CurrentFile_User_Spike1 = 'N/A'

                if User_TimeRanges_Spike2_Enabled:
                    CurrentFile_User_Spike2 = (40 * sum(User_Spike2_RLU_list)) / len(User_Spike2_RLU_list)
                else:
                    CurrentFile_User_Spike2 = 'N/A'

                RLU_DRValues_User_dict[file_clean] = [SampleID] + \
                                                     [Date] + \
                                                     [Test] + \
                                                     [RLUAverage] + \
                                                     [CurrentFile_User_Background1] + \
                                                     [CurrentFile_User_Background2] + \
                                                     [CurrentFile_User_Background3] + \
                                                     [CurrentFile_User_Slope] + \
                                                     [CurrentFile_User_Spike1] + \
                                                     [CurrentFile_User_Spike2]

            RLU_DRValues_Current_dict[file_clean] = [SampleID] + \
                                                    [Date] + \
                                                    [Test] + \
                                                    [RLUAverage] + \
                                                    [CurrentFile_Current_Background1] + \
                                                    [CurrentFile_Current_Background2] + \
                                                    [CurrentFile_Current_Background3] + \
                                                    [CurrentFile_Current_Slope] + \
                                                    [CurrentFile_Current_Spike1] + \
                                                    [CurrentFile_Current_Spike2]



#  Create dataframes from dictionaries
RLU_DRValues_Current_df = pd.DataFrame.from_dict(RLU_DRValues_Current_dict, "index", columns=['SampleID',
                                                                                              'Date',
                                                                                              'Test',
                                                                                              'RLUAverage',
                                                                                              'Background1',
                                                                                              'Background2',
                                                                                              'Background3',
                                                                                              'Slope',
                                                                                              'Spike1',
                                                                                              'Spike2'])

#  Convert specific columns to numeric
RLU_DRValues_Current_df['RLUAverage'] = pd.to_numeric(RLU_DRValues_Current_df['RLUAverage'], errors='ignore')
RLU_DRValues_Current_df['Background1'] = pd.to_numeric(RLU_DRValues_Current_df['Background1'], errors='ignore')
RLU_DRValues_Current_df['Background2'] = pd.to_numeric(RLU_DRValues_Current_df['Background2'], errors='ignore')
RLU_DRValues_Current_df['Background3'] = pd.to_numeric(RLU_DRValues_Current_df['Background3'], errors='ignore')
RLU_DRValues_Current_df['Slope'] = pd.to_numeric(RLU_DRValues_Current_df['Slope'], errors='ignore')
RLU_DRValues_Current_df['Spike1'] = pd.to_numeric(RLU_DRValues_Current_df['Spike1'], errors='ignore')
RLU_DRValues_Current_df['Spike2'] = pd.to_numeric(RLU_DRValues_Current_df['Spike2'], errors='ignore')

#  Convert date columns to datetime, then to float to allow for excel date formatting
RLU_DRValues_Current_df['Date'] = pd.to_datetime(RLU_DRValues_Current_df['Date'], format='%m/%d/%Y', errors='coerce')
RLU_DRValues_Current_df['Date'] = DateTime_to_ExcelDate(RLU_DRValues_Current_df['Date'])

#  Get number of columns and rows to create excel tables
RLU_DRValues_Current_rows = RLU_DRValues_Current_df.shape[0]
RLU_DRValues_Current_columns = RLU_DRValues_Current_df.shape[1]

if User_TimeRanges:
    RLU_DRValues_User_df = pd.DataFrame.from_dict(RLU_DRValues_User_dict, "index", columns=['SampleID',
                                                                                            'Date',
                                                                                            'Test',
                                                                                            'RLUAverage',
                                                                                            'Background1',
                                                                                            'Background2',
                                                                                            'Background3',
                                                                                            'Slope',
                                                                                            'Spike1',
                                                                                            'Spike2'])

    #  Convert specific columns to numeric
    RLU_DRValues_User_df['RLUAverage'] = pd.to_numeric(RLU_DRValues_User_df['RLUAverage'], errors='ignore')
    RLU_DRValues_User_df['Background1'] = pd.to_numeric(RLU_DRValues_User_df['Background1'], errors='ignore')
    RLU_DRValues_User_df['Background2'] = pd.to_numeric(RLU_DRValues_User_df['Background2'], errors='ignore')
    RLU_DRValues_User_df['Background3'] = pd.to_numeric(RLU_DRValues_User_df['Background3'], errors='ignore')
    RLU_DRValues_User_df['Slope'] = pd.to_numeric(RLU_DRValues_User_df['Slope'], errors='ignore')
    RLU_DRValues_User_df['Spike1'] = pd.to_numeric(RLU_DRValues_User_df['Spike1'], errors='ignore')
    RLU_DRValues_User_df['Spike2'] = pd.to_numeric(RLU_DRValues_User_df['Spike2'], errors='ignore')

    #  Convert date columns to datetime, then to float to allow for excel date formatting
    RLU_DRValues_User_df['Date'] = pd.to_datetime(RLU_DRValues_User_df['Date'], format='%m/%d/%Y',errors='coerce')
    RLU_DRValues_User_df['Date'] = DateTime_to_ExcelDate(RLU_DRValues_User_df['Date'])

#  Write dataframes to excel file
#  Create excel file
with pd.ExcelWriter(Summary_File_Folder + "\\" +
                    'RLU_DR_Summary_' + timestr + ".xlsx", engine='xlsxwriter') as writer:
    RLU_DRValues_Current_df.to_excel(writer, sheet_name='Current_Ranges_Summary',startrow= 4, startcol= 2, index=False, header=True)

    wb = writer.book

    Current_Ranges_Summary_ws = writer.sheets['Current_Ranges_Summary']

    #  Create date formats
    Date_Format_MDY = wb.add_format({'num_format': 'mm/dd/yyyy'})

    #  Apply date formats
    Current_Ranges_Summary_ws.set_column('D:D', None, Date_Format_MDY)

    #  Write headers and limit values (starts at 0)
    Current_Ranges_Summary_ws.write(2, 1, "Start Time Index")
    Current_Ranges_Summary_ws.write(2, 6, Current_Background1_Start)
    Current_Ranges_Summary_ws.write(2, 7, Current_Background2_Start)
    Current_Ranges_Summary_ws.write(2, 8, Current_Background3_Start)
    Current_Ranges_Summary_ws.write(2, 9, Current_Slope_Start)
    Current_Ranges_Summary_ws.write(3, 1, "End Time Index")
    Current_Ranges_Summary_ws.write(3, 6, Current_Background1_End)
    Current_Ranges_Summary_ws.write(3, 7, Current_Background2_End)
    Current_Ranges_Summary_ws.write(3, 8, Current_Background3_End)
    Current_Ranges_Summary_ws.write(3, 9, Current_Slope_End)
    Current_Ranges_Summary_ws.write(2, 14, "Current Limits Pass/Fail Check")
    Current_Ranges_Summary_ws.write(3, 13, "Current Limits")
    Current_Ranges_Summary_ws.write(4, 14, "Background1")
    Current_Ranges_Summary_ws.write(4, 15, "Background2")
    Current_Ranges_Summary_ws.write(4, 16, "Background3")
    Current_Ranges_Summary_ws.write(4, 17, "Slope")
    Current_Ranges_Summary_ws.write(3, 14, Current_Background1_UpperLimit)
    Current_Ranges_Summary_ws.write(3, 15, Current_Background2_UpperLimit)
    Current_Ranges_Summary_ws.write(3, 16, Current_Background3_UpperLimit)
    Current_Ranges_Summary_ws.write(3, 17, Current_Slope_LowerLimit)
    Current_Ranges_Summary_ws.write(2, 17, Current_Slope_UpperLimit)

    if User_DRLimits:
        Current_Ranges_Summary_ws.write(2, 20, "User Limits Pass/Fail Check")
        Current_Ranges_Summary_ws.write(3, 19, "User Limits")
        Current_Ranges_Summary_ws.write(4, 20, "Background1")
        Current_Ranges_Summary_ws.write(4, 21, "Background2")
        Current_Ranges_Summary_ws.write(4, 22, "Background3")
        Current_Ranges_Summary_ws.write(4, 23, "Slope")

        if User_DRLimits_Background1_Enabled:
            Current_Ranges_Summary_ws.write(3, 20, User_Background1_UpperLimit)

        if User_DRLimits_Background2_Enabled:
            Current_Ranges_Summary_ws.write(3, 21, User_Background2_UpperLimit)

        if User_DRLimits_Background3_Enabled:
            Current_Ranges_Summary_ws.write(3, 22, User_Background3_UpperLimit)

        if User_DRLimits_Slope_Enabled:
            Current_Ranges_Summary_ws.write(3, 23, User_Slope_LowerLimit)
            Current_Ranges_Summary_ws.write(2, 23, User_Slope_UpperLimit)


    #  Write pass/fail formulas
    for col in range(14,18):
        for row in range(6, 6 + RLU_DRValues_Current_rows):
            if col == 17:
                Current_Ranges_Summary_ws.write_formula(column_number_to_string(col) + str(row),
                                                        '=IF(AND(' + column_number_to_string(col-8) + str(row) + '>' + column_number_to_string(col) + '4,' + column_number_to_string(col-8) + str(row) + '<' + column_number_to_string(col) + '3),"Pass","Fail")')
            else:
                Current_Ranges_Summary_ws.write_formula(column_number_to_string(col) + str(row),
                                               '=IF(' + column_number_to_string(col-8) + str(row) + '>$' + column_number_to_string(col) + '$4,"Fail","Pass")')

    if User_DRLimits:
        for col in range(20, 24):
            for row in range(6, 6 + RLU_DRValues_Current_rows):
                if col == 23:
                    Current_Ranges_Summary_ws.write_formula(column_number_to_string(col) + str(row),
                                                            '=IF(AND(' + column_number_to_string(col - 14) +
                                                            str(row) + '>' + column_number_to_string(col) + '4,'
                                                            + column_number_to_string(col - 14) + str(row) + '<'
                                                            + column_number_to_string(col) + '3),"Pass","Fail")')
                else:
                    Current_Ranges_Summary_ws.write_formula(column_number_to_string(col) + str(row),
                                                            '=IF(' + column_number_to_string(col - 14) + str(
                                                                row) + '>$' + column_number_to_string(
                                                                col) + '$4,"Fail","Pass")')

    #  Repeat sheet creation if user entered time ranges
    if User_TimeRanges:
        RLU_DRValues_User_df.to_excel(writer, sheet_name='User_Ranges_Summary', startrow=4, startcol=2,
                                      index=False, header=True)

        User_Ranges_Summary_ws = writer.sheets['User_Ranges_Summary']

        #  Apply date formats
        User_Ranges_Summary_ws.set_column('D:D', None, Date_Format_MDY)

        #  Write headers and limit values (starts at 0)
        User_Ranges_Summary_ws.write(2, 1, "Start Time Index")

        if User_TimeRanges_Background1_Enabled:
            User_Ranges_Summary_ws.write(2, 6, User_Background1_Start)
            User_Ranges_Summary_ws.write(3, 6, User_Background1_End)

        if User_TimeRanges_Background2_Enabled:
            User_Ranges_Summary_ws.write(2, 7, User_Background2_Start)
            User_Ranges_Summary_ws.write(3, 7, User_Background2_End)

        if User_TimeRanges_Background3_Enabled:
            User_Ranges_Summary_ws.write(2, 8, User_Background3_Start)
            User_Ranges_Summary_ws.write(3, 8, User_Background3_End)

        if User_TimeRanges_Slope_Enabled:
            User_Ranges_Summary_ws.write(2, 9, User_Slope_Start)
            User_Ranges_Summary_ws.write(3, 9, User_Slope_End)

        User_Ranges_Summary_ws.write(3, 1, "End Time Index")
        User_Ranges_Summary_ws.write(2, 14, "Current Limits Pass/Fail Check")
        User_Ranges_Summary_ws.write(3, 13, "Current Limits")
        User_Ranges_Summary_ws.write(4, 14, "Background1")
        User_Ranges_Summary_ws.write(4, 15, "Background2")
        User_Ranges_Summary_ws.write(4, 16, "Background3")
        User_Ranges_Summary_ws.write(4, 17, "Slope")

        User_Ranges_Summary_ws.write(3, 14, Current_Background1_UpperLimit)
        User_Ranges_Summary_ws.write(3, 15, Current_Background2_UpperLimit)
        User_Ranges_Summary_ws.write(3, 16, Current_Background3_UpperLimit)
        User_Ranges_Summary_ws.write(3, 17, Current_Slope_LowerLimit)
        User_Ranges_Summary_ws.write(2, 17, Current_Slope_UpperLimit)

        if User_DRLimits:
            User_Ranges_Summary_ws.write(2, 20, "User Limits Pass/Fail Check")
            User_Ranges_Summary_ws.write(3, 19, "User Limits")
            User_Ranges_Summary_ws.write(4, 20, "Background1")
            User_Ranges_Summary_ws.write(4, 21, "Background2")
            User_Ranges_Summary_ws.write(4, 22, "Background3")
            User_Ranges_Summary_ws.write(4, 23, "Slope")

            if User_DRLimits_Background1_Enabled:
                User_Ranges_Summary_ws.write(3, 20, int(User_Background1_UpperLimit))

            if User_DRLimits_Background2_Enabled:
                User_Ranges_Summary_ws.write(3, 21, int(User_Background2_UpperLimit))

            if User_DRLimits_Background3_Enabled:
                User_Ranges_Summary_ws.write(3, 22, int(User_Background3_UpperLimit))

            if User_DRLimits_Slope_Enabled:
                User_Ranges_Summary_ws.write(3, 23, float(User_Slope_LowerLimit))
                User_Ranges_Summary_ws.write(2, 23, float(User_Slope_UpperLimit))

        #  Write pass/fail formulas
        for col in range(14, 18):
            for row in range(6, 6 + RLU_DRValues_Current_rows):
                if col == 17:
                    User_Ranges_Summary_ws.write_formula(column_number_to_string(col) + str(row),
                                                            '=IF(AND(' + column_number_to_string(col - 8) + str(
                                                                row) + '>' + column_number_to_string(
                                                                col) + '4,' + column_number_to_string(col - 8) + str(
                                                                row) + '<' + column_number_to_string(
                                                                col) + '3),"Pass","Fail")')
                else:
                    User_Ranges_Summary_ws.write_formula(column_number_to_string(col) + str(row),
                                                            '=IF(' + column_number_to_string(col - 8) + str(
                                                                row) + '>$' + column_number_to_string(
                                                                col) + '$4,"Fail","Pass")')

        if User_DRLimits:
            for col in range(20, 24):
                for row in range(6, 6 + RLU_DRValues_Current_rows):
                    if col == 23:
                        User_Ranges_Summary_ws.write_formula(column_number_to_string(col) + str(row),
                                                                '=IF(AND(' + column_number_to_string(col - 14) +
                                                                str(row) + '>' + column_number_to_string(col) + '4,'
                                                                + column_number_to_string(col - 14) + str(row) + '<'
                                                                + column_number_to_string(col) + '3),"Pass","Fail")')
                    else:
                        User_Ranges_Summary_ws.write_formula(column_number_to_string(col) + str(row),
                                                                '=IF(' + column_number_to_string(col - 14) + str(
                                                                    row) + '>$' + column_number_to_string(
                                                                    col) + '$4,"Fail","Pass")')



