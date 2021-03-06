# -*- coding: utf-8 -*-
"""
Created on Sun Feb  7 11:39:18 2021

@author: Jacopo

script to convert youtube links into invidio.us links in an excel file

see this workaround if excel doesn't open the links anymore:
https://docs.microsoft.com/en-us/office/troubleshoot/office-suite-issues/cannot-locate-server-when-click-hyperlink'

"""

import os
import openpyxl

# %% Invidious parameters

# list of public instances: https://instances.invidio.us/
# list of url parameters: https://docs.invidious.io/List-of-URL-parameters.md

base = 'https://invidious.tube/'
dm = 'dark_mode=true'
autoplay = 'autoplay=0'
q = 'quality=hd720'
volume = 'volume=0'
v = str()

# parameters list, add here if you want new ones
par = [v, dm, autoplay, q, volume]


# %%
# set directory and name of your original file, and the name of the transformed file
# note that you need double escape after the disk name (e.g. C:\\...)
directory = "C:\\Users\Jacopo\OneDrive - Universita degli Studi di Milano-Bicocca\Personal\Documents\Allenamento"
name = "WorkOutInvidious.xlsx"
final_name = "WorkOutTry.xlsx"

# set number of rows and columns you want to scan for yt links!
n_row = 80
n_col = 17


# %% transform link function

def transform_link(text):

    '''Takes in a youtube link, transforms it into an invidious one
    with the set parameters'''

    param = par.copy()
    param[0] = text[text.find('v='):text.find('v=')+13]
    if 'watch?' not in text:
        param[0] = 'v=' + text.split('/')[-1][:11]

    if 't=' in text:
        t = text[text.find('t='):]
        if '&' in t:
            t = t.split('&')[0]
        param.append(t)

    new_link = base + 'watch?' +'&' + '&'.join(param)
    return new_link

# just some examples to try out the function:

links = ['https://www.youtube.com/watch?v=yoyzOkB95IQ&t=14s',
          'https://youtube.com/watch?v=yoyzOkB95IQ&t=14s',
          'https://youtu.be/lbozu0DPcYI?t=42s',
          'https://m.youtube.com/watch?v=9iHM6X6uUH8',
          'https://youtu.be/lbozu0DPcYI',
          'https://www.youtube.com/watch?v=yoyzOkB95IQ',
          'https://m.youtube.com/watch?v=9iHM6X6uUH8&t=32s',
          'https://www.youtube.com/watch?t=1m52s&v=vdjWgw98EeI',
          'https://invidious.snopyta.org/watch?t=1m52s&v=vdjWgw98EeI&dark_mode=true&autoplay=0&quality=hd720&volume=0']

for link in links:
    print('\n' + transform_link(link))


# %%

os.chdir(directory)
wb = openpyxl.open(name)


sheets = wb.sheetnames

for name in sheets:
    sheet = wb[name]
    for row in range(1, n_row):
        for col in range(1, n_col):
            cell = sheet.cell(row, col)
            # identify youtube link:
            if cell.hyperlink is not None and ('youtu' in cell.hyperlink.target or
                                               'invidious' in cell.hyperlink.target):
                cell.hyperlink.target = transform_link(cell.hyperlink.target)
                cell.style = "Hyperlink"

            if isinstance(cell.value, str) and (('HYPERLINK' and 'youtu' in cell.value) or
                                                ('HYPERLINK' and 'invidious' in cell.value)):
                cell.hyperlink = transform_link(cell.value.split('"')[1])
                cell.value = cell.value.split('"')[-2]
                cell.style = "Hyperlink"

wb.save(final_name)
