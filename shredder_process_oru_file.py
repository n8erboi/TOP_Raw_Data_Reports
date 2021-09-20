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
import pathlib
import matplotlib.pyplot as plt
import numpy as np
import math

class ORU_File_Readings:
    def __init__(self, id=None):
        """

        :param id: ORU_Channel
        :param readings: Readings for the specific ORU channel
        """

        self.id = id
        self.readings = []

    def append_readings(self, reading):
        """
        Append new ORU read value
        :return:
        """
        self.readings.append(reading)

    def calc_reading_deltas(self):
        """
        Calculate difference between each reading, max delta, and the index of the max delta
        :return:
        """
        self.reading_deltas = np.diff(self.readings)
        self.max_delta = np.max(self.reading_deltas)
        self.max_delta_index = np.where(self.reading_deltas == self.max_delta)[0][0]

        self.calc_result()

    def calc_result(self):
        """
        Calculate the test result
        :return:
        """
        self.curve_deltas = self.reading_deltas[(self.max_delta_index + 10):]
        self.curve_delta_min = np.min(self.curve_deltas)
        self.curve_delta_min_index = np.where(self.curve_deltas == self.curve_delta_min)[0][0]
        self.result = self.curve_delta_min_index / 10

        self.calc_moving_average_result()


    def calc_moving_average_result(self):
        """
        Calculate the moving averages of the readings and the result based on the moving averages
        :return:
        """
        moving_average_n = 5
        self.cumsum = np.cumsum(self.curve_deltas)
        self.cumsum[moving_average_n:] = self.cumsum[moving_average_n:] - self.cumsum[:-moving_average_n]
        self.curve_deltas_moving_average = self.cumsum[moving_average_n - 1:] / moving_average_n
        self.curve_deltas_moving_average_min = np.min(self.curve_deltas_moving_average)
        self.curve_deltas_moving_average_min_index = np.where(self.curve_deltas_moving_average == self.curve_deltas_moving_average_min)[0][0]

        self.moving_average_result = self.curve_deltas_moving_average_min_index / 10

    def calc_mabs_result(self):
        """
        Convert units to mAbs and calculate the result
        :return:
        """

        self.curve_mabs = -np.log(self.readings)*1000






#oru_file_folder = pathlib.Path(easygui.diropenbox("Select The Directory Of ORU Data File(s)"))
oru_file_folder = pathlib.Path(r"\\SYSFILESVR6\rnd\Systems Engineering\Fan4\Bacon")
oru_file_list = list(oru_file_folder.rglob("*.txt"))
oru_obj_list = []

#  Loop through all ORU files
for file in oru_file_list:

    open_file = open(file)

    for line in open_file:

        line = line.rstrip("\n")

        splitline = line.split(" ")

        #  Save info from current line
        current_oru = splitline[3]
        current_channel = splitline[5]
        current_oru_id = current_oru + "_" + current_channel

        #  Save and convert ORU reading from hex to int
        current_reading_hex = splitline[9]
        current_reading_int = int(current_reading_hex, 16)

        #  Create ORU object if one doesn't exist yet for this ORU_Channel
        if not any([current_oru_id == x.id for x in oru_obj_list]):
            current_oru_obj = ORU_File_Readings(id=current_oru_id)
            oru_obj_list.append(current_oru_obj)

        #  Set current ORU object equal to existing object if it already exists
        else:
            bool_list = [current_oru_id == x.id for x in oru_obj_list]
            matching_id_index = [i for i, val in enumerate(bool_list) if val]
            current_oru_obj = oru_obj_list[matching_id_index[0]]

        #  Save data for current ORU_channel
        current_oru_obj.append_readings(current_reading_int)

#  Create plots for each test
for oru in oru_obj_list:

    oru.calc_reading_deltas()
    oru.calc_mabs_result()
    print(oru.curve_mabs)

    figure, ax = plt.subplots(figsize=(12, 4))

    #ax.plot(oru.curve_mabs, marker='o')  #  mAbs
    ax.plot(oru.curve_deltas_moving_average, marker='o')  #  Moving average
    #ax.plot(oru.readings[(oru.max_delta_index + 10):], marker='o') #  Raw data

    figure.suptitle("Curve Deltas, 5-Point Moving Average")
    ax.axvline(x=oru.curve_deltas_moving_average_min_index)
    ax.text(x=.05, y=.05, s="Estimated Result: " + str(oru.moving_average_result) + " s", fontsize=15, transform=ax.transAxes)
    plt.show()

for oru in oru_obj_list:
    figure, ax = plt.subplots(figsize=(12, 4))
    ax.plot(oru.readings)
