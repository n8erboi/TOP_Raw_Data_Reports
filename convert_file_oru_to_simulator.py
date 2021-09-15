import time
import easygui
import os
import pandas as pd
import csv


#  This script converts ORU files into txt files to be read by the AM Simulator

#  Create timestamp for output filename
timestr = time.strftime("%Y_%m_%d-%H%M%S")

#  Raw data files filepath
ORU_File_Folder = easygui.diropenbox("Select The Directory Of ORU Data File(s)")
Sim_File_Folder = easygui.diropenbox("Select The Directory To Save The Simulator File(s) To")

for path, subdirs, files in os.walk(ORU_File_Folder):
    for file in files:

        if ".csv" in file.lower() and file.lower()[0] == 'd':

            filename_remove_csv = file.split('.')[0]

            #  Initiate dictionaries to save AM data for each channel
            All_Channels_AMData_dict = {}

            #  Read each line of file
            for line in open(ORU_File_Folder + '\\' + file, encoding='utf-8'):

                #  Remove new line character at the end of the current line
                line = line.rstrip("\n")

                #  Split line at each comma
                splitline = line.split(',')

                #  Check if line contains ORU data
                if 'ORU' in line and 'Read=' in line:
                    if 'End' in line:
                        continue

                    #  Save channel identifier (length depends on whether date and time values are included in line)
                    if len(splitline) == 7:
                        Channel_ID = str(splitline[4].strip())
                        if len(Channel_ID) == 3:
                            Channel_ID = '0' + Channel_ID
                        AM_Data_Value = int(splitline[5].strip())

                    elif len(splitline) == 5:
                        Channel_ID = int(splitline[2].strip())
                        if len(Channel_ID) == 3:
                            Channel_ID = '0' + Channel_ID
                        AM_Data_Value = int(splitline[3].strip())

                    #  Save AM data value to corresponding channel dictionary
                    if Channel_ID in All_Channels_AMData_dict:
                        All_Channels_AMData_dict[Channel_ID] = All_Channels_AMData_dict[Channel_ID] + [AM_Data_Value]
                    else:
                        All_Channels_AMData_dict[Channel_ID] = [AM_Data_Value]

            for Channel_AMData_list_key in All_Channels_AMData_dict:

                #  Create output csv file for current channel
                current_file = open(Sim_File_Folder + '\\' + filename_remove_csv + '_Channel_' + str(Channel_AMData_list_key) + '_' + timestr + '.txt', 'w', newline='')

                #  Loop through every AM data value for current channel
                for i in range(len(All_Channels_AMData_dict[Channel_AMData_list_key])):
                    current_file.write(str(All_Channels_AMData_dict[Channel_AMData_list_key][i]) + ',0,\n')

                #  Close file
                current_file.close()


