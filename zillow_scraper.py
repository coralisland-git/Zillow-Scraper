import multiprocessing as mp
import os
import queue
import re
import time
from multiprocessing import Pool, cpu_count
from urllib.parse import urljoin
import requests
from bs4 import BeautifulSoup
from lxml import html
import xml.etree.ElementTree as ET
import threading
import csv
from urllib.parse import urlencode
from string import ascii_lowercase, ascii_uppercase
import codecs
import random
import time
from Main_GUI import *
from settings import *

TOTAL_RESULTS = queue.Queue()
TOTAL_CNT = 0
TOTAL_COMPLETED = 0
ZILLOW_LINES = []
ALREADY_ZILLOW_LINES = []
HEADING_EXIST = False
RESULT_FILE = None

global Property_Address_Column
global Property_City_Column
global Property_State_Column
global Property_Zip_Column

Property_Address_Column_index = ascii_lowercase.index(Property_Address_Column.lower())
Property_City_Column_index = ascii_lowercase.index(Property_City_Column.lower())
Property_State_Column_index = ascii_lowercase.index(Property_State_Column.lower())
Property_Zip_Column_index = ascii_lowercase.index(Property_Zip_Column.lower())

result_file_name = 'Output/Zillow_Extracting_Details.csv'
proxy_file_name = 'proxies.txt'
proxies = []
with open(proxy_file_name, 'rb') as text:       
    proxies =  [ "http://" + x.decode("utf-8").strip() for x in text.readlines()]

def insert_row(result_row):
    global RESULT_FILE
    RESULT_FILE.write('"' + '","'.join(result_row) + '"' + "\n")
    RESULT_FILE.flush()

if not os.path.exists('Output'):
    os.makedirs('Output')

result_file_name = 'Output/Zillow_Extracting_Details.csv'
RESULT_FILE = None


def insert_row(result_row):
    global RESULT_FILE
    RESULT_FILE.write('"' + '","'.join(result_row) + '"' + "\n")
    RESULT_FILE.flush()


if not os.path.exists('Output'):
    os.makedirs('Output')

if os.path.isfile(result_file_name):
    out_file = open(result_file_name, 'r', encoding='utf-8')
    csv_reader = list(csv.reader(out_file))

    out_file.close()
    RESULT_FILE = codecs.open(result_file_name, "a", "utf-8")
    if len(csv_reader) == 0:
        RESULT_FILE.write(u'\ufeff')
        HEADING_EXIST = False
    elif len(csv_reader) > 1:
        HEADING_EXIST = True
        for i, line in enumerate(csv_reader):
            if i == 0:
                continue
            ALREADY_ZILLOW_LINES.append(line)
else:
    RESULT_FILE = codecs.open(result_file_name, "a", "utf-8")
    RESULT_FILE.write(u'\ufeff')
    HEADING_EXIST = False


def isAlready(line):
    global Property_Address_Column
    global Property_City_Column
    global Property_State_Column
    global Property_Zip_Column
    if len(ALREADY_ZILLOW_LINES) == 0:
        return False
    else:
        for already_zilow_line in ALREADY_ZILLOW_LINES:
            if line[Property_Address_Column_index] == already_zilow_line[Property_Address_Column_index] \
                    and line[Property_City_Column_index] == already_zilow_line[Property_City_Column_index] \
                    and line[Property_State_Column_index] == already_zilow_line[Property_State_Column_index] \
                    and line[Property_Zip_Column_index] == already_zilow_line[Property_Zip_Column_index]:
                return True
        return False


def get_details(args):
    line = args
    global ZILLOW_API_KEYS

    if isAlready(line):
        return ['ALREADY'] + line
    else:
        if ZILLOW_API_KEYS:
            try:
                global Property_Address_Column_index
                global Property_City_Column_index
                global Property_State_Column_index
                global Property_Zip_Column_index

                address = line[Property_Address_Column_index]
                city = line[Property_City_Column_index]
                state = line[Property_State_Column_index]
                post = line[Property_Zip_Column_index]

                zws_id = ZILLOW_API_KEYS[-1]
                ZILLOW_API_KEYS = [zws_id] + ZILLOW_API_KEYS[:-1]
                params = {
                    'zws-id': zws_id,
                    'address': address,
                    'citystatezip': "-".join([city, state, post]),
                    'rentzestimate': 'true'
                }
                url_params = urlencode(params)
                url = 'http://www.zillow.com/webservice/GetDeepSearchResults.htm?' + url_params
                session = requests.Session()
                r = session.get(url, proxies = {'http': random.choice(proxies)})

                root = ET.fromstring(r.text)

                try:
                    code = root.find('message/code').text
                except:
                    code = ''

                if code == '7':
                    if zws_id in ZILLOW_API_KEYS:
                        ZILLOW_API_KEYS.remove(zws_id)

                if code != '0':
                    return ['STATUS', code] + line

                try:
                    ZPID = root.find('*//zpid').text
                except:
                    ZPID = ''
                try:
                    homedetails = root.find('*//homedetails').text
                except:
                    homedetails = ''
                try:
                    zestimate = root.find('*//zestimate/amount').text
                except:
                    zestimate = ''
                try:
                    zestimate_LOW = root.find('*//zestimate/valuationRange/low').text
                except:
                    zestimate_LOW = ''
                try:
                    zestimate_HIGH = root.find('*//zestimate/valuationRange/high').text
                except:
                    zestimate_HIGH = ''
                try:
                    rentzestimate = root.find('*//rentzestimate/amount').text
                except:
                    rentzestimate = ''
                try:
                    rentzestimate_LOW = root.find('*//rentzestimate/valuationRange/low').text
                except:
                    rentzestimate_LOW = ''
                try:
                    rentzestimate_HIGH = root.find('*//rentzestimate/valuationRange/high').text
                except:
                    rentzestimate_HIGH = ''
                try:
                    zindexValue = root.find('*//localRealEstate/region/zindexValue').text
                except:
                    zindexValue = ''
                try:
                    lastSoldDate = root.find('*//lastSoldDate').text
                except:
                    lastSoldDate = ''
                try:
                    lastSoldPrice = root.find('*//lastSoldPrice').text
                except:
                    lastSoldPrice = ''
                try:
                    bedrooms = root.find('*//bedrooms').text
                except:
                    bedrooms = ''
                try:
                    bathrooms = root.find('*//bathrooms').text
                except:
                    bathrooms = ''
                try:
                    lotSizeSqFt = root.find('*//lotSizeSqFt').text
                except:
                    lotSizeSqFt = ''
                try:
                    totalRooms = root.find('*//totalRooms').text
                except:
                    totalRooms = ''
                try:
                    finishedSqFt = root.find('*//finishedSqFt').text
                except:
                    finishedSqFt = ''
                try:
                    yearBuilt = root.find('*//yearBuilt').text
                except:
                    yearBuilt = ''
                try:
                    taxAssessmentYear = root.find('*//taxAssessmentYear').text
                except:
                    taxAssessmentYear = ''
                try:
                    taxAssessment = root.find('*//taxAssessment').text
                except:
                    taxAssessment = ''
                try:
                    useCode = root.find('*//useCode').text
                except:
                    useCode = ''
                try:
                    STATUS_1 = root.find('message/text').text
                except:
                    STATUS_1 = ''

                url = homedetails + '?fullpage=true'

                headers = {
                    'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
                    'accept-encoding': 'gzip, deflate, br',
                    'accept-language': 'en-US,en;q=0.9',
                    'upgrade-insecure-requests': '1',
                    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.77 Safari/537.36'
                }

                session.cookies.clear()
                r = session.get(url, headers=headers, proxies = {'http': random.choice(proxies)})

                tree = html.fromstring(r.text)

                try:
                    raw_STATUS_2 = tree.xpath('//div[@class="estimates"]//div[@class="status"]//text()')
                    STATUS_2 = " ".join(raw_STATUS_2).strip()
                except:
                    STATUS_2 = ''

                try:
                    raw_Price = tree.xpath('//div[@class="estimates"]//div[@class="price"]//text()')
                    Price = " ".join(raw_Price).strip()
                except:
                    Price = ''

                if ':' in STATUS_2 and Price == '':
                    tmp_STATUS_2 = STATUS_2
                    STATUS_2 = tmp_STATUS_2.split(':')[0].strip()
                    Price = tmp_STATUS_2.split(':')[1].strip()

                result_row = line + [
                    ZPID, homedetails, zestimate, zestimate_LOW, zestimate_HIGH, rentzestimate, rentzestimate_LOW,
                    rentzestimate_HIGH,
                    zindexValue, lastSoldDate, lastSoldPrice, bedrooms, bathrooms, lotSizeSqFt, totalRooms, finishedSqFt,
                    yearBuilt,
                    taxAssessmentYear, taxAssessment, useCode, STATUS_1, STATUS_2, Price
                ]

                result_row = [str(elm).strip() if elm else '' for elm in result_row]
                return ['SUCCESS'] + result_row

            except:
                return ['FAIL'] + line
        else:
            return ['ZWSID_EXCEEDED'] + line


class MainScraper():
    def __init__(self, master):
        self.master = master
        self.result_row = queue.Queue()
        self.gui = MainGUI(master, self.result_row)
        self.start_url = "http://www.zillow.com/webservice/GetDeepSearchResults.htm"
        self.running = 0

        self.session_proxies = {'http': 'socks5://127.0.0.1:9150',  # the default port for tor
                                'https': 'socks5://127.0.0.1:9150'}

        manager = mp.Manager()
        self.q = manager.Queue()
        cpu_cnt = cpu_count()
        self.pool = mp.Pool(cpu_cnt)
        # self.pool = mp.Pool(1)
        print("Your computer has {} cores.".format(cpu_cnt))

        self.periodicCall()

    def start_requests(self):
        jobs = []
        global ZILLOW_LINES

        for i, zillow_line in enumerate(ZILLOW_LINES):
            if i == 0:
                global HEADING
                global HEADING_EXIST
                HEADING = zillow_line + [
                    'ZPID', 'homedetails', 'zestimate', 'zestimate LOW', 'zestimate HIGH', 'rentzestimate',
                    'rentzestimate LOW',
                    'rentzestimate HIGH', 'zindexValue', 'lastSoldDate', 'lastSoldPrice ', 'bedrooms', 'bathrooms',
                    'lotSizeSqFt',
                    'totalRooms', 'finishedSqFt', 'yearBuilt', 'taxAssessmentYear', 'taxAssessment', 'useCode',
                    'STATUS', 'STATUS',
                    'Price'
                ]

                if HEADING_EXIST == False:
                    insert_row(result_row=HEADING)
                    HEADING_EXIST = True
                continue
            job = self.pool.apply_async(get_details, (zillow_line,))
            jobs.append(job)
        # collect results from the workers through the pool result queue
        total_completed = 0
        for job in jobs:
            result = job.get()
            if result[0] == 'SUCCESS':
                total_completed += 1
                insert_row(result_row=result[1:])
                msg = str(result)
                self.result_row.put([total_completed, msg])

            elif result[0] == 'FAIL':
                msg = str(result)
                self.result_row.put([total_completed, msg])

                zillow_line = result[1:]
                job = self.pool.apply_async(get_details, (zillow_line,))
                jobs.append(job)
            elif 'STATUS' == result[0]:

                global STATUS_CODE_LIST

                status_code = result[1]
                msg = 'Status Code {}: '.format(status_code) + STATUS_CODE_LIST[status_code] + ', ' + str(result[2:])
                self.result_row.put([total_completed, msg])

                if result[1] == '7':
                    zillow_line = result[2:]
                    job = self.pool.apply_async(get_details, (zillow_line,))
                    jobs.append(job)
                else:
                    total_completed += 1
                    result = result[2:] + [""]*30
                    # global HEADING

                    total_column_cnt = len(HEADING)
                    result = result[:total_column_cnt]
                    result[total_column_cnt-3] = STATUS_CODE_LIST[status_code]
                    insert_row(result_row=result)

            elif result[0] == 'ALREADY':
                total_completed += 1
                insert_row(result_row=result[1:])
                msg = str(result)
                self.result_row.put([total_completed, msg])

            elif result[0] == 'ZWSID_EXCEEDED':
                msg = 'All ZWSIDs Are Exceeded The Limit For The Day.'
                self.result_row.put([total_completed, msg])
                break

        # now we are done, kill the listener
        self.q.put('kill')
        self.pool.close()

    def periodicCall(self):
        self.gui.update_prog_bar()
        if self.gui.running == 1:
            global ZILLOW_LINES
            ZILLOW_LINES = self.gui.zillow_lines
            thread = threading.Thread(target=self.start_requests)
            thread.start()
            self.gui.running = 2
        self.master.after(500, self.periodicCall)

    def endApplication(self):
        self.running = 0


if __name__ == '__main__':
    root = Tk()
    client = MainScraper(root)
    root.mainloop()
