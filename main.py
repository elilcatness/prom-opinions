import os
import sys
import time
from datetime import date as dt, timedelta

from excel_writer import ExcelWriter
from utils import get_json_from_filename, get_driver, get_doc, convert_str_to_dt


def parse_product(url, headers=None):
    doc = get_doc(url, headers=headers)
    presence = doc.xpath('//span[@data-qaid="product_presence"]/span/text()')[0].strip()
    return {
        'Цена': int(doc.xpath('//span[@data-qaid="product_price"]/@data-qaprice')[0]),
        'Наличие': presence
    }


def parse_market(url, driver=None, headers=None):
    if not driver:
        driver = get_driver('binary/chromedriver.exe', headers.get('User-Agent'))
    driver.get(url)
    date_to_check = dt.today() - timedelta(days=365)
    output = []
    starts_translate = {'Отлично': 5,
                        'Хорошо': 4,
                        'Нормально': 3,
                        'Так себе': 2,
                        'Плохо': 1}
    opinions_offset = 0
    while True:
        opinions = driver.find_elements_by_xpath('//li[@data-qaid="opinion_lists"]')
        if not opinions:
            sys.exit(f'Failed to parse opinions {url}')
        start_time = time.time()
        while len(opinions) <= opinions_offset and time.time() - start_time < 20:
            opinions = driver.find_elements_by_xpath('//li[@data-qaid="opinion_lists"]')
        if len(opinions) <= opinions_offset:
            sys.exit(f'Failed to parse new opinions {url}')
        opinions = opinions[opinions_offset:]
        opinions_offset += len(opinions)
        for opinion in opinions:
            header_data = opinion.find_elements_by_xpath(
                './/div[@class="tJcYO _1uX4I"]//div[@class="_2ernF"]//span')
            if not header_data or len(header_data) < 2:
                sys.exit(f'Failed to parse header {url}')
            date_str = header_data[1].text
            if not date_str:
                sys.exit(f'Failed to parse date of opinion {url}')
            date = convert_str_to_dt(date_str)
            if date < date_to_check:
                return output
            data = {
                'Дата': date_str,
                'Звезд': starts_translate.get(opinion.find_element_by_xpath(
                    './/*[@class="_1KcTA _2gVBb _3TpPX"]'
                    '//span[@class="_3h93n"]').text),
                'Товары': []
            }
            products = opinion.find_elements_by_xpath('.//a[@class="_2KaCs _3y8C0 fxzyv"]')
            for product in products:
                product_link = product.get_attribute('href')
                product_data = parse_product(product_link, headers)
                if product_data:
                    data['Товары'].append({'Название': product.text, 'Ссылка': product_link,
                                           **product_data})
            print(', '.join([f'{item[0]}: {item[1]}' for item in data.items()]))
            output.append(data)
        button = driver.execute_script("return document.getElementsByClassName('VspXp _3Cxan _3dQ3K yB0BO')[0];")
        if not button:
            break
        driver.execute_script('arguments[0].click()', button)
    return output


def main(input_filename='input.txt', headers_filename='headers.json', output_filename='output.xlsx'):
    headers = get_json_from_filename(headers_filename)
    if isinstance(headers, str):
        return headers, -1
    if not os.path.exists(input_filename) or not os.path.isfile(input_filename):
        return f'Invalid input_filename {input_filename}', -1
    excel_headers = {}
    links = []
    with open(input_filename, encoding='utf-8') as f:
        for link in f.readlines():
            link = link.strip()
            links.append(link)
            excel_headers[link.split('/')[-1]] = ['Дата', 'Звезд', 'Название', 'Ссылка', 'Цена', 'Наличие']
    writer = ExcelWriter(output_filename, list(excel_headers.keys()), headers=excel_headers)
    writer.write_headers(bold=True)
    for link in links:
        print(f'{"#" * 20} STARTED PARSING {link} {"#" * 20}')
        data = parse_market(link, headers=headers)
        for opinion in data:
            if opinion['Товары']:
                products = opinion.pop('Товары')
                writer.write_row({**opinion, **products[0]})
                if len(products) > 1:
                    for prod in products[1:]:
                        writer.write_row(prod)
            else:
                opinion.pop('Товары')
                writer.write_row(opinion)
    if links:
        print()
    return 'OK'


if __name__ == '__main__':
    callback_data = main()
    if callback_data:
        callback, exit_code = (callback_data if isinstance(callback_data, tuple)
                               and len(callback_data) > 1 else (callback_data, 0))
        print(callback)
        if exit_code != 0:
            sys.exit(exit_code)
