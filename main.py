# ID CLIENT pgjgwxpokkqeg4sp1jcxffyeawaepc
# TOKEN jwuxby25t6avtfmst0ko7ppfyevtyx
# secret = h9xkw35o2vgi0jvfoku2sf1x3fk25o
import csv
import datetime
import os.path
import time
from datetime import datetime
from datetime import date
from tempfile import NamedTemporaryFile

import openpyxl
import requests
from numpy import reshape
from oauth2client.service_account import ServiceAccountCredentials
from openpyxl import Workbook, load_workbook, workbook
from openpyxl.cell import Cell
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.descriptors import Integer, Sequence, String
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.reader.excel import load_workbook
from openpyxl.utils import get_column_letter, quote_sheetname
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.workbook import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from func import Twitch_Auth, col2num, getline, split, total_fr

if __name__ == '__main__':
    # Filename generation
    today = date.today()
    try:
        if not os.path.exists('XLfiles'):
            os.makedirs('XLfiles')
    except OSError:
        print('Error: Creating directory. ' + 'XLfiles')

    filepath = f"StatTwitch{today}"
    cfilepath = f"StatTwitch{today}.xlsx"


    # Creates an excel file
    iter_path = 1
    while os.path.isfile(cfilepath):
        cfilepath = filepath + '_0' + str(iter_path) + '.xlsx'
        iter_path += 1
    os.chdir('XLfiles')
    wb = openpyxl.Workbook()
    wb.guess_types = True
    ws = wb.active
    wb.save(filename=cfilepath)

    # Variables settings
    totalpda = float(0)
    n = 1
    i = 0
    t = 'B'
    col = 1
    firstpassage = True
    line = 1
    NB = 15
    REFRESH = 60

    while n < 50:
        totalviewers = 0

        # Getting datas from Twitch API
        stream = Twitch_Auth()
        stream_data = stream.json()
        stats_stream_fr = 0
        stats_stream_fr = total_fr(stream_data)
        total_frbis = stats_stream_fr
        now = datetime.now()
        today = now.strftime("%H:%M:%S")
        i = 0
        totalpda = float(0)

        # Analizes the TOP_NB Twitch streamers
        while i <= NB:
            pda = (stream_data['data'][i]['viewer_count'] * 100) / total_frbis
            pda = "{: .2f}".format(pda)
            username = stream_data['data'][i]['user_name']
            viewer_count = stream_data['data'][i]['viewer_count']
            title = stream_data['data'][i]['title']
            data = [username, pda]
            # Table format
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
        # Reset every x seconds
        time.sleep(REFRESH)



    wb.save(cfilepath)
