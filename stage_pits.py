# standard
import glob
import os
import shutil
from pathlib import Path
import datetime
import re
import pandas as pd
import dateutil.parser as dparser

# megan's dictionary of 2021 SnowEx time series sites
from site_dict import dict1

'''
Author: Megan Mason
Created: August, 2021
Modified:

This script stages the SnowEx 2021 data that were uploaded to NSIDC.
Data include Snow Pits, Interval Boards, and Depth Transects.

It imports one unquie modeule that stores a python dictionary of the site names and codes.
It organizes all snow pits (and soon to be other data) in to the repected folder for qa/qc purposes.
'''

#------------------------------------------------------------------------------#
# SET-UP
basepath = '/Users/meganmason491/Downloads/data/' #third download from NSIDC, 9/7/2021
destination_folder = '/Users/meganmason491/Google Drive/SnowEx-2021/SnowEx-2021-campaign-core/SnowEx2021_TimeSeries_Pits'

fpaths = []
for path in Path(basepath).rglob('*.xlsx'): #walks through all sub-dirs and looks for xlsx's
    fpaths.append(path)

#------------------------------------------------------------------------------#
# ORGANIZE by MEAUREMENT TYPE
pts = [x for x in fpaths if "SNOW_PIT" in x.name] # pits
ibs = [x for x in fpaths if "INTERVAL_BOARD" in x.name] # interval boards
txs = [x for x in fpaths if "Depth" in x.name] # depth transects

#------------------------------------------------------------------------------#
# NESTED LOOPS - organize into subfolders based on k,v pairs - raw data
print('...sorting raw data files')

for k, v in dict1.items():

    if not os.path.exists(destination_folder + '/nsidc-raw/' + v):
        os.makedirs(os.path.join(destination_folder + '/nsidc-raw/' + v))

    for pt in pts:

        if pt.name.split('_')[0] == k:
            print(f".....file: {pt.name}")
            shutil.copy(pt, destination_folder + '/nsidc-raw/' + v)

#------------------------------------------------------------------------------#
# NESTED LOOPS - organize into subfolders based on k,v pairs - pre format
print('...sorting pre-formated data files')

for k, v in dict1.items():

    if not os.path.exists(destination_folder + '/pre-format/' + v):
        os.makedirs(os.path.join(destination_folder + '/pre-format/' + v))

    for pt in pts:
        if '(1)' in pt.name: # checked all duplicates and they are trash
            pass
        elif pt.name.split('_')[0] == k:
            print(f".....file: {pt.name}")
            shutil.copy(pt, destination_folder + '/pre-format/' + v)


print('complete')
#------------------------------------------------------------------------------#
