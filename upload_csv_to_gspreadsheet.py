#!/usr/bin/env python
# -*- coding: utf-8 -*-

##
## upload_data.py
##
## Upload data to Google drive
##

from __future__ import print_function

import os
import sys
import getopt
from getpass import getpass
import csv
from itertools import chain
import gdata.docs.service
import gdata.docs.client
import gdata.spreadsheet.service

##
def exit(msg):
  print(msg, file=sys.stderr)
  sys.exit(1)

##
def usage():
  print('''\
  Usage: ''' + os.path.basename(sys.argv[0]) +
  ''' worksheet_name:tsv_file ...
     -u/--user         Set Google_ID
     -s/--spreadsheet  Google spreadsheet name
     -w/--worksheet    Default worksheet where data are uploaded
     -h/--help         Show help''')
  sys.exit(0)

##
def get_tsv_data(tsv_file):
  data = []
  msg  = ''

  try:
    reader = csv.reader(open(tsv_file, 'r'), delimiter='\t')
    for row in reader:
      data.append(list(row))
  except ValueError as e:
    msg = 'Not a TSV file'
  except IOError as e:
    msg = "Can not read file `{0}'".format(tsv_file)
 
  return msg, data

##
def upload_tsv_data(gs_client, spread_id, worksheet, data):
  # Justify the size of all data columns
  max_col = max([ len(x) for x in data ])
  [ x.extend([''] * (max_col - len(x))) for x in data ]

  # Search worksheet
  query = gdata.spreadsheet.service.DocumentQuery()
  query.title = worksheet
  query.title_exact = 'true' 
  w_sheet = gs_client.GetWorksheetsFeed(spread_id, query=query)

  if w_sheet.entry:
    gs_client.DeleteWorksheet(worksheet_entry=w_sheet.entry[0])

  # Create new worksheet
  w_sheet_id = gs_client.AddWorksheet(worksheet,
                                      len(data), max_col,
                                      spread_id).id.text.split('/')[-1]

  # Load the TSV data to the worksheet
  query = gdata.spreadsheet.service.CellQuery()
  query.return_empty = 'true'
  cells = gs_client.GetCellsFeed(spread_id, w_sheet_id, query=query)
  batch_request = gdata.spreadsheet.SpreadsheetsCellsFeed()
  data = list(chain.from_iterable(data))

  for i, cv in enumerate(data):
    cells.entry[i].cell.inputValue = cv
    batch_request.AddUpdate(cells.entry[i])

  updated = gs_client.ExecuteBatch(batch_request, cells.GetBatchLink().href)

##
def upload_data(user, passwd, tsv_data, spreadsheet, worksheet):

  # Login to Google drive
  gd_client = gdata.docs.service.DocsService()
  gs_client = gdata.spreadsheet.service.SpreadsheetsService()
  gd_client.ssl = True
  gs_client.ssl = True
  source    = 'UploadData'

  try:
    gd_client.ClientLogin(user, passwd, source=source)
  except gdata.service.BadAuthentication as e:
    exit(str(e) + '.\nAuthentication failed.')
  except Exception as e:
    exit(str(e) + '.\nLogin Error.')

  gs_client.ClientLogin(user, passwd, source=source)

  # Get spreadsheet feed
  query = gdata.spreadsheet.service.DocumentQuery()
  query.title = spreadsheet
  query.title_exact = 'true'
  feed = gs_client.GetSpreadsheetsFeed(query=query)

  if not feed.entry:
    exit('None spreadsheets.')

  spread_id = feed.entry[0].id.text.split('/')[-1]

  # Enter operation mode
  auth_token = gd_client.GetClientLoginToken()
  gd_client.SetClientLoginToken(gs_client.GetClientLoginToken())

  # Upload
  for worksheet, tsv_file in tsv_data:
    msg, data = get_tsv_data(tsv_file)
    if msg:
      print(msg + "\nSkip uploading `{0}'.".format(tsv_file))
      continue
    print("Uploading `{0}' to `{1}:{2}' ...".
          format(tsv_file, spreadsheet, worksheet))
    upload_tsv_data(gs_client, spread_id, worksheet, data)

  # Close operation
  gd_client.SetClientLoginToken(auth_token)

##
def main():

  # Parse arguments
  try:
    opts, args = getopt.getopt(sys.argv[1:], 'u:s:w:h',
                   ['user=', 'spreadsheet=', 'worksheet=', 'help'])
  except getopt.error, msg:
    usage()

  if [ arg for option, arg in opts
       if option == '--help' or option == '-h']:
    usage()

  # TSV files
  if not args:
    print('You should specify a TSV file.\n')
    usage()
  else:
    tsv_file = args

  # Google spreadsheet name
  spreadsheet = [ arg for option, arg in opts
             if option == '--spreadsheet' or option == '-s']
  if spreadsheet:
    spreadsheet = spreadsheet[0]
  else:
    print('You should specify a spreadsheet name.\n')
    usage()

  # Worksheet name
  worksheet = [ arg for option, arg in opts
                if option == '--worksheet' or option == '-w']
  if worksheet:
    worksheet = worksheet[0]

  # Get TSV file
  tsv_data = []
  for f in tsv_file:
    if ':' in f:
      ws, file = f.split(':', 1)
    else:
      if worksheet:
        ws, file = worksheet, f
      else:
        print('You should specify a default worksheet name.\n')
        usage()

    if not ws:
      print('You should specify a worksheet name.\n')
      usage()
    if not file:
      print('You should specify a TSV file.\n')
      usage()

    tsv_data.append([ws, file])

  if not tsv_data:
    print('You should specify a TSV file.\n')
    usage()

  # Google ID
  user = [ arg for option, arg in opts
           if option == '--user' or option == '-u']
  if not user:
    user = raw_input('Please enter your Google ID: ')
  else:
    user = user[0]

  # Password
  passwd = getpass('Enter password: ')

  ## Upload
  upload_data(user, passwd, tsv_data, spreadsheet, worksheet)

if __name__ == '__main__':
  main()
