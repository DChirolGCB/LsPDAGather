''' from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
import requests
from lxml import html
from bs4 import BeautifulSoup
import lxml.html
from selenium import webdriver
driver = webdriver.Chrome


def get_totalviewers():

    Service.start()
    options = Options()
    options.headless = True
    driver = webdriver.Chrome(options=options)
    driver.get('https://twitchtracker.com/languages/French.html')
    print(driver.page_source)
 '''
