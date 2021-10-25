# ID CLIENT pgjgwxpokkqeg4sp1jcxffyeawaepc
# TOKEN jwuxby25t6avtfmst0ko7ppfyevtyx
# secret = h9xkw35o2vgi0jvfoku2sf1x3fk25o
from pyparsing import Word, alphas
from func import total_fr
from func import Twitch_Auth
from func import getline
from func import barchartrace
from func import col2num
from func import split
#from webscrap import get_totalviewers
import openpyxl
import csv
import datetime
import os.path
import csv
import time
from openpyxl.chart import BarChart, Series, Reference
import gspread

from openpyxl import workbook
from numpy import reshape
import bar_chart_race as bcr
from openpyxl.workbook import Workbook
from openpyxl import load_workbook
import pandas as pd
import csv
# import matplotlib.pyplot as plt
# import pandas as pd
import requests
import re
import time
from datetime import datetime

import unittest
from tempfile import NamedTemporaryFile
import gspread
import requests
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.cell import Cell
import bar_chart_race as bcr
from openpyxl.descriptors import (
    String,
    Sequence,
    Integer,
)
import string
import copy
import sys
import warnings
import re
import requests
from oauth2client.service_account import ServiceAccountCredentials
import re
import time
import unittest
from lxml import html
from tempfile import NamedTemporaryFile
import gspread
import requests
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.cell import Cell
import bar_chart_race as bcr
from openpyxl.descriptors import (
    String,
    Sequence,
    Integer,
)
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.reader.excel import load_workbook
from openpyxl.styles import Color, Fill
from openpyxl.styles import numbers
from openpyxl.utils import quote_sheetname
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.workbook import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.reader.excel import load_workbook
from openpyxl.styles import Color, Fill
from openpyxl.styles import numbers
from openpyxl.utils import quote_sheetname
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from datetime import date
import matplotlib.pyplot as plt


# get_totalviewers()
# exit()
plt.rcParams['animation.ffmpeg_path'] = "C:/FFmpeg/bin/ffmpeg"

today = date.today()
try:
    if not os.path.exists('XLfiles'):
        os.makedirs('XLfiles')
except OSError:
    print('Error: Creating directory. ' + 'XLfiles')

filepath = f"StatTwitch{today}"
cfilepath = f"StatTwitch{today}.xlsx"
print(filepath)
# import matplotlib.pyplot as plt
# import pandas as pd
# import win32com.client
# import xlwings as xw


# Création excel
iter_path = 1
while os.path.isfile(cfilepath):
    cfilepath = filepath + '_0' + str(iter_path) + '.xlsx'
    iter_path += 1
os.chdir('XLfiles')
wb = openpyxl.Workbook()
wb.guess_types = True
ws = wb.active
wb.save(filename=cfilepath)


totalpda = float(0)
n = 1
i = 0
t = 'B'
col = 1
firstpassage = True
line = 1
# if n == 1:
#    exit()
# barchartrace('sample.xlsx')

while n < 50:
    totalviewers = 0

    stream = Twitch_Auth()
    stream_data = stream.json()
    stats_stream_fr = 0
    stats_stream_fr = total_fr(stream_data)
    total_frbis = stats_stream_fr
    now = datetime.now()
    today = now.strftime("%H:%M:%S")
    i = 0
    totalpda = float(0)
    while i <= 15:

        pda = (stream_data['data'][i]['viewer_count'] * 100) / total_frbis
        pda = "{: .2f}".format(pda)
        username = stream_data['data'][i]['user_name']
        viewer_count = stream_data['data'][i]['viewer_count']
        title = stream_data['data'][i]['title']
        data = [username, pda]
        if firstpassage:
            ws[f'A{2}'] = today
            ws[f'{t}1'] = username
            ws[f'{t}2'] = float(pda)
        else:
            try:
                t = get_column_letter(getline(username, ws))
            except:
                t = getline(username, ws)
            ws[f'A{line}'] = today
            ws[f"{t}1"] = username
            ws[f"{t}{line}"] = float(pda)
        totalpda = "{: .2f}".format(totalpda)
        totalpda = float(totalpda) + float(pda)
        t = col2num(t)
        t = get_column_letter(t + 1)
        i += 1
    line += 1
    wb.save("sample.xlsx")
    firstpassage = False
    wb.save(cfilepath)
    time.sleep(2)
# print(totalpda)


totalpda = "{: .2f}".format(totalpda)
print(
    f"Les 20 premiers streamers fr représentent {totalpda}% des vues du top 100")

wb.save(cfilepath)
