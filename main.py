# ID CLIENT pgjgwxpokkqeg4sp1jcxffyeawaepc
# TOKEN jwuxby25t6avtfmst0ko7ppfyevtyx
# secret = h9xkw35o2vgi0jvfoku2sf1x3fk25o
print("test")
import csv
import datetime
import os.path
import re
import time
import unittest
from tempfile import NamedTemporaryFile

import gspread
# import matplotlib.pyplot as plt
# import pandas as pd
import requests
# import win32com.client
# import xlwings as xw
from oauth2client.service_account import ServiceAccountCredentials
from openpyxl import Workbook
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl import load_workbook
from openpyxl.cell import Cell
from openpyxl.descriptors import (
    String,
    Sequence,
    Integer,
)
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.reader.excel import load_workbook
from openpyxl.styles import Color, Fill
from openpyxl.styles import numbers
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.workbook import Workbook

from func import Twitch_Auth
from func import total_fr

# Création excel

wb = load_workbook(filename='test.xlsx')
wb.guess_types = True
ws = wb.active
# Connexion Twitch


# Calcul Total Viewer Fr


# Calcul PDA Top100

totalpda = float(0)
n = 1
i = 0

while n < 50:
    stream = Twitch_Auth()
    stream_data = stream.json()
    stats_stream_fr = 0
    stats_stream_fr = total_fr(stream_data)
    total_frbis = stats_stream_fr
    i = 0
    totalpda = float(0)
    while i <= 9:
        pda = (stream_data['data'][i]['viewer_count'] * 100) / total_frbis
        pda = "{: .2f}".format(pda)
        username = stream_data['data'][i]['user_name']
        viewer_count = stream_data['data'][i]['viewer_count']
        title = stream_data['data'][i]['title']
        data = [username, pda]
        ws[f'B{i+1+n}'] = username
        ws[f'C{i+1+n}'] = float(pda)
        ws[f'D{i+1+n}'] = viewer_count
        ws[f'E{i+1+n}'] = title
        totalpda = "{: .2f}".format(totalpda)
        totalpda = float(totalpda) + float(pda)
        i += 1
    wb.save("sample.xlsx")
    time.sleep(15)
    n += 11

print(totalpda)

totalpda = "{: .2f}".format(totalpda)

print(f"Les 100 premiers streamers fr représentent {totalpda}% des vues")

