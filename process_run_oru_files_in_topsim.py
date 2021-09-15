import time
import easygui
import os
import math
import pyautogui as pag
import pygetwindow as pgw
import win32gui
import win32com.client
import pandas as pd
import csv


#  This script converts ORU files into txt files to be read by the AM Simulator

#  Create timestamp for output filename
timestr = time.strftime("%Y_%m_%d-%H%M%S")

#  Set 1 second delay after each pyautogui command
pag.PAUSE = 0.5

#  TOP SW coordinate positions
sample_rack_xy = [[618, 214],
                  [735, 214],
                  [855, 214],
                  [972, 214],
                  [1090, 214],
                  [1204, 214],
                  [1327, 214],
                  [1440, 214]]

sample_rack_next_right_xy = [1665, 92]

sample_test_selection = [[991, 220],
                         [991, 280],
                         [991, 350],
                         [991, 415],
                         [991, 475],
                         [991, 540],
                         [991, 605],
                         [991, 733]]

test_selection_window_tabs = [[1490, 190],
                              [1600, 190],
                              [1720, 190]]

test_selection_grid = [[1520, 270],
                       [1660, 270],
                       [1820, 270],
                       [1520, 350],
                       [1660, 350],
                       [1820, 350],
                       [1520, 430],
                       [1660, 430],
                       [1820, 430],
                       [1520, 505],
                       [1660, 505],
                       [1820, 505]]

test_selection_grid_close = [1740, 640]

run_icon = [140, 90]

temp_out_confirm_run_start = [920, 596]

menu_selection_analysis = [240, 40]
menu_selection_analysis_sample_area = [290, 120]


#  Raw data files filepath
oru_file_folder = easygui.diropenbox("Select The Directory Of ORU Data File(s)")
sim_file_folder = easygui.diropenbox("Select The Directory To Save The Simulator File(s) To")

run_user_details_msg = "Please enter the run details\n\n" \
                       "SampleID:\nNumbering will be appended by this script starting at 000\n" \
                       "NC_Run1 > NC_Run1_000, NC_Run1_001...\n\n" \
                       "Test Code:\nNumbering will be appended by this script starting at 00\n" \
                       "PTRead > PTRead00, PTRead01..."
run_user_details_field_names = ['SampleID (max 12 characters)', 'Test Code (max 8 characters)', 'AM Data Points per Test']
run_user_details_values = easygui.multenterbox(msg=run_user_details_msg,
                                               title='Run Details',
                                               fields=run_user_details_field_names)

#  Run details error checking
while 1:
    if run_user_details_values is None:
        break
    error_msg = ''

    am_data_is_a_number = True

    for i in range(len(run_user_details_values)):
        if run_user_details_values[i].strip == "":
            error_msg = error_msg + ('"%s" is a required field.\n\n' % run_user_details_field_names[i])
        elif i == 0 and len(run_user_details_values[i]) > 12:
            error_msg = error_msg + ('"%s" must be 12 characters or less.\n\n' % run_user_details_field_names[i])
        elif i == 1 and len(run_user_details_values[i]) > 8:
            error_msg = error_msg + ('"%s" must be 8 characters or less.\n\n' % run_user_details_field_names[i])
        if i == 2:
            try:
                int(run_user_details_values[i])
            except ValueError:
                error_msg = error_msg + ('"%s" must be an integer.\n\n' % run_user_details_field_names[i])
                am_data_is_a_number = False
        if i == 2 and am_data_is_a_number and not isinstance(int(run_user_details_values[i]), int):
            error_msg = error_msg + ('"%s" must be an integer.\n\n' % run_user_details_field_names[i])

    #  No problems found
    if error_msg == '':
        break
    run_user_details_values = easygui.multenterbox(msg=error_msg,
                                                   title='Run Details',
                                                   fields=run_user_details_field_names,
                                                   values=run_user_details_values)

user_sampleid = run_user_details_values[0]
user_testcode = run_user_details_values[1]
user_amdata_points = int(run_user_details_values[2])

confirm_sim_rack_removal = easygui.msgbox("Confirm all racks are removed from the instrument in the ACL TOP Simulator", "Rack Removal Check", "OK")

#  Create list of sample ids
sample_id_list = []
for i in range(0, 120):
    if i < 10:
        sample_id_list.append(user_sampleid + '_00' + str(i))
    elif 10 <= i < 100:
        sample_id_list.append(user_sampleid + '_0' + str(i))
    else:
        sample_id_list.append(user_sampleid + '_' + str(i))

#  Create list of test ids
test_list = []
for i in range(0, 30):
    if i < 10:
        test_list.append(user_testcode + '0' + str(i))
    else:
        test_list.append(user_testcode + str(i))


am_data_files_list = []
raw_data_file_counter = 0

for path, subdirs, files in os.walk(oru_file_folder):
    for file in files:

        if ".csv" in file.lower() and file.lower()[0] == 'd':

            filename_remove_csv = file.split('.')[0]

            #  Initiate dictionaries to save AM data for each channel
            all_channels_amdata_dict = {}

            #  Read each line of file
            for line in open(oru_file_folder + '\\' + file, encoding='utf-8'):

                #  Remove new line character at the end of the current line
                line = line.rstrip("\n")

                #  Split line at each comma
                splitline = line.split(',')

                #  Check if line contains ORU data
                if 'ORU' in line and 'Read=' in line:
                    if 'End' in line:
                        continue

                    #  Save channel identifier (length depends on whether date and time values are included in line)
                    #  NEED TO WRITE HEADER TO FILE
                    if len(splitline) == 7:
                        channel_id = str(splitline[4].strip())
                        if len(channel_id) == 3:
                            channel_id = '0' + channel_id
                        am_data_value = int(splitline[5].strip())

                    elif len(splitline) == 5:
                        channel_id = int(splitline[2].strip())
                        if len(channel_id) == 3:
                            channel_id = '0' + channel_id
                        am_data_value = int(splitline[3].strip())

                    #  Save AM data value to corresponding channel dictionary
                    if channel_id in all_channels_amdata_dict:
                        all_channels_amdata_dict[channel_id] = all_channels_amdata_dict[channel_id] + [am_data_value]
                    else:
                        all_channels_amdata_dict[channel_id] = [am_data_value]

            for channel_amdata_list_key in all_channels_amdata_dict:

                #  Create output csv file for current channel
                current_filepath = sim_file_folder + '\\' + filename_remove_csv + '_Channel_' + str(channel_amdata_list_key) + '_' + timestr + '.txt'
                current_file = open(sim_file_folder + '\\' + filename_remove_csv + '_Channel_' + str(channel_amdata_list_key) + '_' + timestr + '.txt', 'w', newline='')
                am_data_files_list.append(filename_remove_csv + '_Channel_' + str(channel_amdata_list_key) + '_' + timestr + '.txt')

                current_file.write("//\n")
                #current_file.write(sample_id_list[math.floor(raw_data_file_counter/30)] + ",19,05,2021,13,57,00," + str(test_list[raw_data_file_counter % 30]) + ",0,0,0,\n")
                current_file.write(sample_id_list[math.floor(raw_data_file_counter/28)] + ",19,05,2021,13,57,00," + str(test_list[raw_data_file_counter % 28]) + ",0,0,0,\n")
                current_file.write("//\n")
                current_file.write("0,0,1,0,0,1,\n")
                current_file.write("//\n")
                current_file.write("1,1000,0,0,0,\n")
                current_file.write("//\n")


                #  Loop through every AM data value for current channel, add/skip am data points if needed
                am_data_line_counter = 0
                for i in range(len(all_channels_amdata_dict[channel_amdata_list_key])):
                    if am_data_line_counter == user_amdata_points:
                        continue
                    current_file.write(str(all_channels_amdata_dict[channel_amdata_list_key][i]) + ',0,\n')

                    am_data_line_counter += 1

                if len(all_channels_amdata_dict[channel_amdata_list_key]) < user_amdata_points:
                    missing_data_points = user_amdata_points - len(all_channels_amdata_dict[channel_amdata_list_key])
                    for i in range(0, missing_data_points):
                        current_file.write(str(all_channels_amdata_dict[channel_amdata_list_key][len(all_channels_amdata_dict[channel_amdata_list_key]) - 1]) + ',0,\n')

                #  Close file
                raw_data_file_counter += 1
                current_file.close()

#
#  Create simulator script file to load materials and raw data files
#

test_total = raw_data_file_counter
sample_total = math.ceil(test_total / 28)
sample_rack_total = sample_total // 10 + 1
sim_file_name = 'Script-File_' + timestr + '.act'

sample_test_total_dict = {}

#  Track number of tests for each sample
for i in range(0, sample_total):

    if test_total - 28 - i * 28 > 0:
        current_sample_total_tests = 28
    else:
        current_sample_total_tests = test_total - i * 28

    sample_test_total_dict[i] = current_sample_total_tests

with open(sim_file_folder + "\\" + sim_file_name, "w") as openfile:

    #  Set sample containers. Loop through each sample rack, sample position, and test
    for i in range(0, sample_rack_total):

        #  Determine last vial used in current sample rack
        if i == sample_rack_total - 1:
            end_vial = sample_total % 10
        else:
            end_vial = 10

        for j in range(0, end_vial):

            sample_id_counter = 10 * i + j

            openfile.write("<ACT>12 SetVialLabel</ACT>\n")
            openfile.write("<AREA>1</AREA>\n")
            openfile.write("<TRACK>" + str(i) + "</TRACK>\n")
            openfile.write("<VIAL>" + str(j) + "</VIAL>\n")
            openfile.write("<LABEL>" + str(len(sample_id_list[sample_id_counter])) + ' ' + sample_id_list[sample_id_counter] + "</LABEL>\n\n")

        #  Set rack label
        openfile.write("<ACT>12 SetRackLabel</ACT>\n")
        openfile.write("<AREA>1</AREA>\n")
        openfile.write("<TRACK>" + str(i) + "</TRACK>\n")
        openfile.write("<LABEL>2 S" + str(i + 1) + "</LABEL>\n\n")

        #  Move barcode reader to current rack
        openfile.write("<ACT>10 MoveReader</ACT>\n")
        openfile.write("<AREA>1</AREA>\n")
        openfile.write("<TRACK>" + str(i) + "</TRACK>\n\n")

        #  Wait for barcode reader, insert rack, wait for rack
        openfile.write("<ACT>13 WaitForReader</ACT>\n\n")
        openfile.write("<ACT>23 TryToInsertOrRemoveRack</ACT>\n\n")
        openfile.write("<ACT>11 WaitForRack</ACT>\n\n")

    #  Insert PT-Readiplastin
    openfile.write("<ACT>12 SetVialLabel</ACT>\n")
    openfile.write("<AREA>2</AREA>\n")
    openfile.write("<TRACK>0</TRACK>\n")
    openfile.write("<VIAL>0</VIAL>\n")
    openfile.write("<LABEL>12 008500000000</LABEL>\n\n")

    #  Insert Clean B Diluted
    openfile.write("<ACT>12 SetVialLabel</ACT>\n")
    openfile.write("<AREA>2</AREA>\n")
    openfile.write("<TRACK>0</TRACK>\n")
    openfile.write("<VIAL>1</VIAL>\n")
    openfile.write("<LABEL>12 099500000000</LABEL>\n\n")

    #  Set rack label
    openfile.write("<ACT>12 SetRackLabel</ACT>\n")
    openfile.write("<AREA>2</AREA>\n")
    openfile.write("<TRACK>0</TRACK>\n")
    openfile.write("<LABEL>2 R1</LABEL>\n\n")

    #  Move barcode reader to current rack
    openfile.write("<ACT>10 MoveReader</ACT>\n")
    openfile.write("<AREA>2</AREA>\n")
    openfile.write("<TRACK>0</TRACK>\n")

    #  Wait for barcode reader, insert rack, wait for rack
    openfile.write("<ACT>13 WaitForReader</ACT>\n")
    openfile.write("<ACT>23 TryToInsertOrRemoveRack</ACT>\n")
    openfile.write("<ACT>11 WaitForRack</ACT>\n\n")

    #  Load all raw data files
    for i in range(len(am_data_files_list)):

        #corrected_filepath = "\\" + am_data_files_list[i].replace("\\\\", "\\")
        sim_file_single_bs = sim_file_folder.replace("\\\\", "\\") + "\\" + am_data_files_list[i]

        if sim_file_single_bs[0] == "\\":
            sim_file_single_bs = "\\" + sim_file_single_bs

        openfile.write("<ACT>11 LoadRawData</ACT>\n")
        openfile.write("<FILE>" + str(len(sim_file_single_bs)) + " " + sim_file_single_bs + "</FILE>\n\n")


    openfile.close()


#
#  Program tests using pyautogui and click run
#

#  Get handles for TOP SIM and TOP SW
all_window_titles = pgw.getAllTitles()

for title in all_window_titles:
    if "AMS" in title and "IP" in title:
        acl_top_sim_title = title
        acl_top_sim_window = pgw.getWindowsWithTitle(acl_top_sim_title)[0]
    elif "ACL TOP" in title and "AMSim" in title:
        acl_top_sw_title = title
        acl_top_sw_window = pgw.getWindowsWithTitle(acl_top_sw_title)[0]


#  Activate TOP SIM, set to full screen
acl_top_sim_window.activate()
acl_top_sim_window.maximize()
pag.sleep(3)

#  Run script file (Menu -> Run Script File)
pag.hotkey('alt', 'm')
pag.press('down')
pag.press('enter')
pag.sleep(1)

#  Use tab key to select filepath textbox
pag.press('tab', presses=5)
pag.press('enter')
pag.sleep(1)
pag.press('enter')
pag.sleep(1)

#  Type SIM script filepath
pag.write(sim_file_folder)

pag.press('enter')
pag.sleep(1)
pag.write(sim_file_name)
pag.press('enter')
pag.sleep(1)

#  Wait for script to run
pag.sleep(10)

#  Acitvate TOP SW, set to full screen
acl_top_sw_window.activate()
acl_top_sw_window.maximize()
pag.sleep(3)

#  Select Sample Screen (Analysis -> Sample Screen)
pag.click(x=menu_selection_analysis[0], y=menu_selection_analysis[1])
pag.sleep(1)
pag.click(x=menu_selection_analysis_sample_area[0], y=menu_selection_analysis_sample_area[1])
pag.sleep(2)

#  Loop through each sample rack
for i in range(0, sample_rack_total):

    #  Select current sample rack (click sample rack icon if first rack, click next rack button if not first rack)
    if i == 0:
        pag.click(x=sample_rack_xy[i][0], y=sample_rack_xy[i][1], clicks=2)
    else:
        pag.click(x=sample_rack_next_right_xy[0], y=sample_rack_next_right_xy[1], clicks=2)

    #  Determine last vial used in current sample rack
    if i == sample_rack_total - 1:
        end_vial = sample_total % 10
    else:
        end_vial = 10

    #  Loop through each sample ID
    for j in range(0, end_vial):

        sample_id_counter = 10 * i + j

        #  Click current sample test selection area
        pag.click(x=sample_test_selection[j][0], y=sample_test_selection[j][1] )

        #  Loop through each test and click
        for k in range(0, sample_test_total_dict[sample_id_counter]):

            #  Click test tab if needed
            if k == 0:
                pag.click(x=test_selection_window_tabs[0][0], y=test_selection_window_tabs[0][1])
            elif k == 12:
                pag.click(x=test_selection_window_tabs[1][0], y=test_selection_window_tabs[1][1])
            elif k == 24:
                pag.click(x=test_selection_window_tabs[2][0], y=test_selection_window_tabs[2][1])

            #  Click current test
            pag.click(x=test_selection_grid[k % 12][0], y=test_selection_grid[k % 12][1])

#  Close test selection window
pag.click(x=test_selection_grid_close[0], y=test_selection_grid_close[1])

#  Run tests
pag.click(x=run_icon[0], y=run_icon[1])
pag.click(x=temp_out_confirm_run_start[0], y=temp_out_confirm_run_start[1])



