# ID CLIENT pgjgwxpokkqeg4sp1jcxffyeawaepc
# TOKEN jwuxby25t6avtfmst0ko7ppfyevtyx
# secret = h9xkw35o2vgi0jvfoku2sf1x3fk25o

from datetime import datetime as dateuser
import string
from pyparsing import Word, alphas
import time
from openpyxl.utils import get_column_letter
import requests
from oauth2client.service_account import ServiceAccountCredentials
from openpyxl.cell import Cell
from openpyxl.descriptors import (
    String,
    Sequence,
    Integer,
)
from openpyxl.workbook import Workbook


def Twitch_Auth():
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
             "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        'F:\ProjetsPython\StreamStatFinalM\secrets.json', scope)
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
