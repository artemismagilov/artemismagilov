from oauth2client.service_account import ServiceAccountCredentials
import apiclient
import httplib2
import argparse
import csv
import os
import sys

parser = argparse.ArgumentParser(epilog="""
                         instruction: add the arguments in strict sequence - sheet name→ range data→ secret key→ 
                         id sheet in link→ filename(optional).\n
                         example: [path]/python3 spreadsheet.py Sheet1 A1:C7 secret_key.json 273839dss3dad033901 
                         [path/]report.[txt|csv](optional)
                         """)
parser.add_argument("name_sheet", help="name sheet in Google Sheets")
parser.add_argument("range_data", help="specify the data range in the sheet in Google Sheets")
parser.add_argument("secret_key", help="the name of the downloaded file with the private key")
parser.add_argument("id_sheet", help="the id name from your table link")
parser.add_argument("--file_name", "-f", help="file name with its extension (optional)", default=0)
args = parser.parse_args()
if os.path.exists(args.secret_key):
    credentials = ServiceAccountCredentials.from_json_keyfile_name(args.secret_key,
                                                                   ['https://www.googleapis.com/auth/spreadsheets',
                                                                    'https://www.googleapis.com/auth/drive'])
else:
    print(f'A non-existent secret_key was passed {args.secret_key}. The file must be in json format')
    quit()
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)
range_name = args.name_sheet + '!' + args.range_data
try:
    table = service.spreadsheets().values().get(spreadsheetId=args.id_sheet, range=range_name).execute()
except:
    print(f'Invalid values passed {args.id_sheet} {range_name}. range_name must be in the format - A1:C7 or '
          f'wrong id_sheet. The link may be restricted in access '
          )
    quit()
lines = table.get('values', None)
if not lines:
    print(f'There is no such sheet {args.name_sheet}')
    quit()
strings = ''
for line in lines:
    strings += '\t'.join(line) + '\n'
if os.path.isfile(args.file_name):
    with open(file=args.file_name, mode='w', encoding='utf-8') as f:
        filename, file_extension = os.path.splitext(args.file_name)
        if file_extension == '.csv':
            csv_writer = csv.writer(f)
            for line in lines:
                csv_writer.writerow(line)
        elif file_extension == '.txt':
            f.write(strings)
else:
    print('This file does not exist')
    sys.stdout.write(strings)
