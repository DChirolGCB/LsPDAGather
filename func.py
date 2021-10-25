# ID CLIENT pgjgwxpokkqeg4sp1jcxffyeawaepc
# TOKEN jwuxby25t6avtfmst0ko7ppfyevtyx
# secret = h9xkw35o2vgi0jvfoku2sf1x3fk25o
from datetime import datetime as dateuser
import datetime
import warnings
import sys
import copy
import string
from pyparsing import Word, alphas
import csv
import datetime
import os.path
import csv
import time
from openpyxl import Workbook
import gspread
from openpyxl.utils import get_column_letter
import openpyxl
from numpy import reshape
import bar_chart_race as bcr
from openpyxl.workbook import Workbook
from openpyxl import load_workbook
import pandas as pd
import csv
# import matplotlib.pyplot as plt
# import pandas as pd
import requests
from oauth2client.service_account import ServiceAccountCredentials
import re
import time
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
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.reader.excel import load_workbook
from openpyxl.styles import Color, Fill
from openpyxl.styles import numbers
from openpyxl.utils import quote_sheetname
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.workbook import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
import matplotlib.pyplot as plt
plt.rcParams['animation.ffmpeg_path'] = "C:/FFmpeg/bin/ffmpeg"


def Twitch_Auth():
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
             "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        'F:\ProjetsPython\StreamStatFinalM\secrets.json', scope)
    client = gspread.authorize(credentials)
    client_id = 'pgjgwxpokkqeg4sp1jcxffyeawaepc'
    client_secret = 'h9xkw35o2vgi0jvfoku2sf1x3fk25o'
    starttime = time.time()

    body = {
        'client_id': client_id,
        'client_secret': client_secret,
        "grant_type": 'client_credentials'
    }
    r = requests.post('https://id.twitch.tv/oauth2/token', body)
    keys = r.json()
    headers = {
        'Client-ID': client_id,
        'Authorization': 'Bearer ' + keys['access_token']
    }
    header = ['streamer_login', 'viewer_count', 'PDA', 'title']
    stream = requests.get(
        'https://api.twitch.tv/helix/streams?first=100&language=fr', headers=headers)
    return stream


def split(word):
    return [char for char in word]


def col2num(col):
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num


def total_fr(stream_data):
    (stream_data)
    total_fr = 0
    i = 0
    pda = 0
    totalpda = 0
    while i < 98:
        total_fr += stream_data['data'][i]['viewer_count']
        i += 1

    now = dateuser.now()
    today = now.strftime("%H:%M:%S")
    print(f"{today} : Il y a {total_fr} viewers fr sur les 100 premiers streams fr")
    return total_fr


def getline(username, ws):
    for columns in ws.iter_cols(1):
        for cell in columns:
            if cell.value == username:
                return(cell.column)
    return get_column_letter(len(ws['1']) + 1)


def barchartrace(source_race):
    df = pd.DataFrame(pd.read_excel(source_race))
    bcr.bar_chart_race(
        df=df,
        filename='testvideo.mp4',
        orientation='h',
        sort='desc',
        n_bars=20,
        fixed_order=False,
        fixed_max=True,
        steps_per_period=20,
        period_length=1000,
        interpolate_period=False,
        period_label={'x': .98, 'y': .3, 'ha': 'right', 'va': 'center'},
        period_summary_func=lambda v, r: {'x': .98, 'y': .2,
                                          's': f'Total deaths: {v.sum():,.0f}',
                                          'ha': 'right', 'size': 11},
        perpendicular_bar_func='median',
        title='PDA jour du',
        bar_size=.95,
        shared_fontdict=None,
        scale='linear',
        fig=None,
        writer=None,
        bar_kwargs={'alpha': .7},
        filter_column_colors=False)
