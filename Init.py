from googleapiclient.discovery import build
from google.oauth2 import service_account
from Func_GSAPI import *

 #Define o escopo e as credenciais do bot
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'credenciais.json'

#Define o ID da planilha
SAMPLE_SPREADSHEET_ID = "15ftB4dwmhDC4eeBR0Redw-BSuqS2elypdFX1I5DauI8"

def Init(spreadsheet , sheet_id):
  
  requests = []

  requests.append(change_size_row(sheet_id , 1 , 2 , 100))

  requests.append(change_size_column(sheet_id , 0 , 1 , 210))
  requests.append(change_size_column(sheet_id , 1 , 2 , 80))
  requests.append(change_size_column(sheet_id , 2 , 3 , 53))
  requests.append(change_size_column(sheet_id , 3 , 4 , 80))
  requests.append(change_size_column(sheet_id , 4 , 5 , 215))
  requests.append(change_size_column(sheet_id , 5 , 6 , 215))
  requests.append(change_size_column(sheet_id , 6 , 7 , 25))

  requests.append(merge_rows(sheet_id , 1 , 4 , 2 , 6))

  requests.append(add_value(sheet_id , 1 , 0 , "Contest name"))
  requests.append(add_value(sheet_id , 1 , 1 , "Editorial"))
  requests.append(add_value(sheet_id , 1 , 2 , "Code"))
  requests.append(add_value(sheet_id , 1 , 3 , "Date"))
  requests.append(add_value(sheet_id , 1 , 4 , "Subject(s)"))

  requests.append(centralize_cells(sheet_id , 1 ,  0 , 2 , 6))
  requests.append(border_all_cells(sheet_id , 1 ,  0 , 2 , 6))
  requests.append(bold_cells(sheet_id , 1 ,  0 , 2 , 6))


  

  body = {"requests":requests}
  spreadsheet.batchUpdate(spreadsheetId=SAMPLE_SPREADSHEET_ID, body=body).execute()

def main():
  creds = None
  creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

  service = build('sheets', 'v4', credentials=creds)
  spreadsheet = service.spreadsheets()

  Init(spreadsheet,0)

if __name__ == '__main__':
  main()
