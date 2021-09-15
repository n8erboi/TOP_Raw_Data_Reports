import time
import easygui
import os
import pandas as pd
import numpy as np
import io
from datetime import datetime
import matplotlib.pyplot as plt

#  Read csv file into pandas
RLU_Data_df = pd.read_csv(filepath_or_buffer="C:\\Users\\jrollett\\Desktop\\Biokit Backups\\BioKit_Pilots_RLUData.csv",
                          sep=",")

#plt.figure()

#RLU_Data_df['RLU'].plot()

RLU_Data_B2G_df = RLU_Data_df.loc[RLU_Data_df['Test'] == 'B2G']
RLU_Data_CAL1_df = RLU_Data_df.loc[RLU_Data_df['Material'] == 'CAL1']

RLU_Data_B2G_CAL1_df = RLU_Data_df.loc[(RLU_Data_df['Material'] == 'CAL1') & (RLU_Data_df['Test'] == 'B2G')]

RLU_Data_Grouped_df = RLU_Data_B2G_CAL1_df.groupby(['Reagent_Cartridge_Lot', 'Reagent_Cartridge_SN']).agg({'RLU': ['mean', 'std']})
RLU_Data_Grouped_df.columns = ['RLU_mean', 'RLU_std']
RLU_Data_Grouped_df = RLU_Data_Grouped_df.reset_index()

#RLU_Data_Grouped_df['RLU_mean'].plot()

#  Overall Boxplot
#boxplot = RLU_Data_B2G_CAL1_df.boxplot(column='RLU', by=['Reagent_Cartridge_Lot', 'Reagent_Cartridge_SN'])

#  Subplots by reagent lot, plot each cartridge sn within each lot
boxplot = RLU_Data_B2G_CAL1_df.groupby(['Reagent_Cartridge_Lot']).boxplot(column='RLU', by='Reagent_Cartridge_SN')
plt.xticks(rotation=90)

plt.show()

#print('stop')

#  Read each line of file
#for line in open("C:\\Users\\jrollett\\Desktop\\Biokit Backups\\BioKit_Pilots_RLUData.csv", encoding='utf-8'):
