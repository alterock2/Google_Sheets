import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials

import gspread
import pandas as pd
from df2gspread import df2gspread as d2g

credentials_file ='stable-woods-374912-3063b446a841.json'
credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file,
['https://www.googleapis.com/auth/spreadsheets',
'https://www.googleapis.com/auth/drive'])
#httpAuth = credentials.authorize(httplib2.Http())
#service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)




gc = gspread.authorize(credentials)

spreadsheet_key = '1s9sQyX9j4rQKMiTErmjkc8r1pp-yxXA5ehMUOun9gpw'
wks_name = 'Master'

df = pd.read_excel(r"C:\Users\user\Desktop\Сервис.xlsx")



d2g.upload(df, spreadsheet_key, wks_name, credentials=credentials, row_names=True, col_names=True)