#!/usr/bin/env python
# -*- coding: utf-8 -*-

##
## download_sample_sheet.py
##
## Download a spreadsheet from Google drive to STDOUT
##

from __future__ import print_function

import os
import sys
import getopt
from getpass import getpass
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
  Usage: ''' + os.path.basename(sys.argv[0]) + '''sheet_name
     -u/--user    Set Google_ID
     -o/--output  Output file(TSV format)
     -h/--help    Show help''')
  sys.exit(0)

##
def download_google_spread_sheet(user, passwd, sheet_name, output):

  ## Login to Google drive
  gd_client = gdata.docs.service.DocsService()

  gs_client = gdata.spreadsheet.service.SpreadsheetsService()
  gd_client.ssl = True
  gs_client.ssl = True
  source    = 'DownloadGoogleSpreadSheet'

  try:
    gd_client.ClientLogin(user, passwd, source=source)
  except gdata.service.BadAuthentication as e:
    exit(str(e) + '.\nAuthentication failed.')
  except Exception as e:
    exit(str(e) + '.\nLogin Error.')

  gs_client.ClientLogin(user, passwd, source=source)

  ## Query
  title = 'title=' + sheet_name
  query = gdata.docs.service.DocumentQuery(text_query=title)
  feed  = gd_client.Query(query.ToUri())

  if not feed.entry:
    exit('None documents.')

  ## Download spreadsheet
  auth_token = gd_client.GetClientLoginToken()
  gd_client.SetClientLoginToken(gs_client.GetClientLoginToken())
  gd_client.Download(feed.entry[0].resourceId.text, output, 'tsv', gid=0)
  gd_client.SetClientLoginToken(auth_token)

##
def main():

  ## Parse arguments
  try:
    opts, args = getopt.getopt(sys.argv[1:], 'u:o:h', ['user=','output=','help'])
  except getopt.error, msg:
    usage()

  if [ arg for option, arg in opts
       if option == '--help' or option == '-h']:
    usage()

  # Google spreadsheet name
  if not args:
    print('You should specify Google sheet name.\n')
    usage()
  else:
    sheet_name = args[0]

  # Output file
  output = [ arg for option, arg in opts
             if option == '--output' or option == '-o']
  if output:
    output = output[0]
  else:
    output = '/dev/stdout'

  # Google ID
  user = [ arg for option, arg in opts
           if option == '--user' or option == '-u']
  if not user:
    user = raw_input('Please enter your Google ID: ')
  else:
    user = user[0]

  # Password
  passwd = getpass('Enter password: ')

  ## Download
  download_google_spread_sheet(user, passwd, sheet_name, output)

if __name__ == '__main__':
  main()
