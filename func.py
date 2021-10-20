import csv
import time
from openpyxl import Workbook
import gspread
import matplotlib.pyplot as plt
import pandas as pd
import requests
from oauth2client.service_account import ServiceAccountCredentials


def Twitch_Auth():
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
             "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        'secrets.json', scope)
    client = gspread.authorize(credentials)
    client_id = 'pgjgwxpokkqeg4sp1jcxffyeawaepc'
    client_secret = '6j62oi625sjrij4btfi3re0r6rql8f'
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
    stream = requests.get('https://api.twitch.tv/helix/streams?first=100&language=fr', headers=headers)
    print(header)
    return stream


def total_fr(stream_data):
    total_fr = 0
    i = 0

    while i < 98:
        total_fr += stream_data['data'][i]['viewer_count']
        i += 1

    print(f"Il y a {total_fr} viewers fr")
    return total_fr
