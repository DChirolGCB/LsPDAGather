# ID CLIENT pgjgwxpokkqeg4sp1jcxffyeawaepc
# TOKEN jwuxby25t6avtfmst0ko7ppfyevtyx
# secret = 6j62oi625sjrij4btfi3re0r6rql8f
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import numbers
import csv
import datetime
import time
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import gspread
import matplotlib.pyplot as plt
import pandas as pd
import requests
from oauth2client.service_account import ServiceAccountCredentials
from openpyxl import load_workbook
from openpyxl.reader.excel import load_workbook
from openpyxl.workbook import Workbook
from openpyxl.styles import Color, Fill
from openpyxl.cell import Cell
from func import Twitch_Auth
from func import total_fr
from tempfile import NamedTemporaryFile
import re
import xlwings as xw
from openpyxl.descriptors import (
    String,
    Sequence,
    Integer,
)
from openpyxl.descriptors.serialisable import Serialisable
import unittest
import os.path
import win32com.client


# Création excel
wb = load_workbook(filename='sample.xlsx')
wb.guess_types = True
ws = wb.active
# Connexion Twitch
stream = Twitch_Auth()
stream_data = stream.json()

# Calcul Total Viewer Fr
total_fr = total_fr(stream_data)
# Calcul PDA Top100
totalpda = 0

n = 0
y = 0
while y < 2:
    i = 0
    while i <= 9:
        pda = (stream_data['data'][i]['viewer_count'] * 100) / total_fr
        totalpda += pda
        pda = "{: .2f}".format(pda)
        pda = pda.replace('.', ',').replace(' ', '')
        username = stream_data['data'][i]['user_name']
        viewer_count = stream_data['data'][i]['viewer_count']
        title = stream_data['data'][i]['title']
        data = [username, pda]
        print(type(pda))
        ws.merge_cells(f'A{n+1}', f'A{n+11}')
        ws[f'B{i+1+n}'] = username
        ws[f'C{i+1+n}'] = viewer_count
        ws[f'D{i+1+n}'] = pda
        ws[f'E{i+1+n}'] = title
        i += 1
    n += 10
    y += 1
ws.insert_rows(1)
totalpda = "{: .2f}".format(totalpda)

print(f"Les 100 premiers streamers fr représentent {totalpda}% des vues")
wb.save("sample.xlsx")
