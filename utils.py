import json
import os.path
import sys
from json import JSONDecodeError

import requests
from lxml import html
from selenium.webdriver import Chrome, ChromeOptions
from datetime import date as dt


def get_json_from_filename(filename: str):
    if not os.path.exists(filename) and not os.path.isfile(filename):
        return f'Invalid filename {filename}'
    try:
        with open(filename, encoding='utf-8') as f:
            return json.loads(f.read())
    except JSONDecodeError:
        return f'Failed to decode JSON {filename}'


def convert_str_to_dt(raw_str):
    try:
        day, month, year = map(int, raw_str.split('.'))
        return dt(year=year, month=month, day=day)
    except ValueError:
        return None


def get_doc(url, params=None, headers=None):
    response = requests.get(url, params=params, headers=headers)
    if not response:
        with open('index.html', 'w', encoding='utf-8') as f:
            f.write(response.text)
        sys.exit('YOU PROXY MAZAFAKA')
    return html.fromstring(response.text)


def get_driver(path='', user_agent=None):
    options = ChromeOptions()
    options.add_argument('--headless')
    if user_agent and isinstance(user_agent, str):
        options.add_argument(f'user-agent={user_agent}')
    return Chrome(path, options=options)