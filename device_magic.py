# -*- coding: utf-8 -*-
"""
Created on Thu Jan 18 14:38:42 2018

@author: hmamin
"""

# -*- coding: utf-8 -*-
"""
Created on Wed Jun 14 14:14:17 2017

@author: hmamin
"""

import sys
from backports import ssl
import imapclient
import pprint
import pyzmail
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
import re
import gspread, time
from oauth2client.service_account import ServiceAccountCredentials
from auto_email import quick_mail


def extract_auth_info(filename):
    '''Read gmail auth info from file and return as list.'''
    with open(filename, 'r') as f:
        txt = f.read()
        return txt.split(',')
    
    
def login(email, password):
    '''Authorize and login to gmail account.'''
    context = ssl.SSLContext(ssl.PROTOCOL_TLSv1_2)
    imapObj = imapclient.IMAPClient('imap.gmail.com', ssl=True, ssl_context=context)
    imapObj.login(email, password)
    imapObj.select_folder('INBOX', readonly=True)
    return imapObj


def parse_raw_emails(newSites, text):
    '''Parse imapObj.fetch object into PyzMessage object.'''
    ptext = []
    if len(newSites)>0:
        for i in newSites:
            ptext.append(pyzmail.PyzMessage.factory(text[i][b'BODY[]']))   #testing loading multiple sites
        return ptext
    else:
        sys.exit('No new sites.')
        
    
def parse_to_text(ptext):
    '''Pass in pyzmessage object, return list of email text.'''
    body = []
    for j in ptext:
        body.append(BeautifulSoup(j.html_part.get_payload().decode(), "lxml").text)
    return body


def auth_gdrive():
    '''Authorize access to Google drive to allow spreadsheet editing.'''
    scope = ['https://spreadsheets.google.com/feeds']
    creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
    client = gspread.authorize(creds)
    return client


def extract_field(field_regex, email_body, site_row, field_name):
    '''Pass in regex to search for in parsed email body, then append to dict
    "cells" in class instance.
    '''
    field_results = re.search(field_regex, email_body)
    try:
        site_row.cells[field_name] = field_results.group().strip()
    except:
        field_results = None
        pass
    

def search_install_list(site):
    '''Find inverter and battery size from install list and assign to 
    relevant dict values for the current site instance.
    '''
    install_cells = install.findall(site.cells['site'])
    if len(install_cells) > 0:
        site.cells['inverterSize'], site.cells['batterySize'] =\
            (cell.value for cell in install.range(install_cells[0].row, 7,\
            install_cells[0].row, 8))
        
        
def find_empty_row(worksheet):
    '''Find empty row to insert new row.'''
    rows = worksheet.row_count
    siteRange = worksheet.range('A110:A{}'.format(rows))
    for i in siteRange:
        if i.value == '':
            max_row = i.row
            break
    print('First available row:', max_row)
    return max_row


class newRow():
    '''New row in MDT for 1 site, from 1 commissioning email.'''
    
    def __init__(self):
        self.cells = {
            'site': '',
            'address': '',
            'inverterSize': '',
            'batterySize': '',
            #'numInverters': '',
            #'numTowers': '',
            'meterNum': '',
            'IP': '',
            'NOCid': '',
            'WirelessNum': '',
            'dateComm': '',
            'fCT': '',
            'gsCT': '',
            'quant': '',
            'zigbee': '',
            'towerMnum': '',
            'towerSnum': '',
            'abID': '',
            'cID': '',
            'IMEI': '',
            'SIM': '',
            'converterNum': '',
            'switchgearNum': '',
            'switchgearSNum': '',
            'acuvim1': '',
            'acuvim2': '',
            'acuvimHard': '',
            'acuvimSoft': '',
            'SysCon': '',
            'DataServer': '',
            'RSSI': '',
            'EDeltas': '',
            'CFtype': '',
            'uploader': '',
            'dev-server': '',
            'controller': '',
            'thermalLogic': '',
            'currentThrottling': '',
            'PVsize': '',
            'systemStorage': '',
            'ESS': '',
            'DPG': '',
            'alerts': '',
            'solar': '',
            'single_inverter': '',
            'single_battery': '',
            'modem': '',
            'erm_mac': ''
            }
        
        
# Select date range to search - now using hard-coded timeframe for task scheduler
#num_days = int(input('Search emails from the past _ days: '))
num_days = 1
endPeriod = datetime.now()
startPeriod = (endPeriod - timedelta(days=num_days)).strftime('%d-%b-%Y')  #time period to search for
print(startPeriod, ' - ', endPeriod.strftime('%d-%b-%Y %H:%M'))

# Login and search emails
#subject_string = 'Controller has been commissioned'
subject_string = 'A new GridSynergy Controller is online:'
user, password = extract_auth_info('gmail_auth.txt')
imapObj = login(user, password)
newSites = imapObj.search(['SUBJECT', subject_string, \
                           'SINCE', startPeriod, 'FROM', 'noreply@devicemagic.com'])
print('newSites: ', newSites)                                                     #helpful to see list of emails found
text = imapObj.fetch(newSites, ['BODY[]'])  #unparsed email body text
ptext = parse_raw_emails(newSites, text)
body = parse_to_text(ptext)
imapObj.logout()

# Access Google Drive and open sheet
client = auth_gdrive()
MDT = client.open('2Copy of MDT').sheet1
#MDT = client.open('40-000055 Master Deployment Tracking').sheet1         #actual mdt - permissions issues       
install = client.open('copy_install').sheet1
          
allRows = [newRow() for i in range(len(newSites))]

# loop through list of new site objects and fill in dict fields from email text
re_patterns = [('(?<=Site Name:) .+', 'site'), 
               ('(?<=Site Address:) .+', 'address'),
               ('(?<=Meter Number:) .+', 'meterNum'), 
               ('(?<=IP:) .+', 'IP'),
               ('(?<=ERM Computer Mac Address:) .+', 'erm_mac'),
               ('(?<=\nComputer MAC Address:) .+', 'abID'), 
               ('(?<=building metering method\?:) .+', 'fCT'),
               ('(?<=KYZ Multiplier\?:) .+', 'quant')]
for k, site_row in enumerate(allRows):
    for pattern in re_patterns:
        extract_field(pattern[0], body[k], site_row, pattern[1])
        
        # Date field requires extra processing
        dateRe = re.search(r'(\d\d\d\d-\d\d-\d\d)', body[k])
        try:
            d = time.strptime(dateRe.group(), "%Y-%m-%d")
        except:
            dateRe = None    
            pass
        site_row.cells['dateComm'] = time.strftime("%m/%d/%Y", d)   

# Find system size from Install List
for site in allRows:
    search_install_list(site)
pprint.pprint([a.cells for a in allRows])       

# Check for each new site in MDT, and add row if no duplicates found.
to_emails = ['hmamin@greencharge.net', 'hmamin@engiestorage.com']
subject = 'PyDeviceMagic finished running.'
msg = ''

# Insert row if not already in MDT
t = time.time()  
for s in allRows:
    print('\nSite:', s.cells['site'])
    try:
        duplicates = MDT.find(s.cells['site'])
    except gspread.CellNotFound:
        max_row = find_empty_row(MDT)
        MDT.insert_row(list(s.cells.values()), index=max_row)
        print(s.cells['site'] + ' added to MDT.')
        subject = f'MDT Updated - New Site Commissioned'
        msg += f'\n{s.cells["site"]} added to MDT\nRow {max_row}\n\n\
        https://docs.google.com/spreadsheets/d/1jMn7z96KnGYCZMuXP56I1l4Q_flmr_BUdE4mbP0QaW0/edit'
    else:
        print(s.cells['site'] + ' already exists in MDT')
        print('Duplicates:', duplicates)
        msg += f"\n{s.cells['site']} already on MDT."
print('\ntime: ', time.time()-t)            #time to append row after finding info
quick_mail(to_emails, subject, msg)