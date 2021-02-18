from googleapiclient.discovery import build
from google.oauth2 import service_account
from Func_GSAPI import *

#Define o escopo e as credenciais do bot
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'credenciais.json'

#Define o ID da planilha
SAMPLE_SPREADSHEET_ID = "15ftB4dwmhDC4eeBR0Redw-BSuqS2elypdFX1I5DauI8" 

num_usuarios = 5

class Contest:
  name , link , editorial_link , code , date , num_question = "" , "" , "" , "" , "" , 0
  def __init__(self , name , link , editorial_link , code , date , num_question):
    self.name = name
    self.link = link 
    self.editorial_link = editorial_link 
    self.code = code
    self.date = date
    self.num_question = num_question

def insert_contest(spreadsheet , sheet_id , contest):
  requests = []
  requests.append(insert_rows(sheet_id,contest.num_question))
  requests.append(merge_columns(sheet_id,0,0,contest.num_question,4))
  requests.append(add_hyperlink(sheet_id,0,0,contest.name,contest.link))
  if contest.editorial_link != "":
    requests.append(add_hyperlink(sheet_id,0,1,"Editorial",contest.editorial_link))
  requests.append(add_condition_rule(sheet_id,0,5,contest.num_question,5+num_usuarios))
  requests.append(add_value(sheet_id,0,2,contest.code))
  requests.append(add_value(sheet_id,0,3,contest.date))
  requests.append(centralize_cells(sheet_id,0,0,contest.num_question,5+num_usuarios))
  requests.append(border_all_cells(sheet_id,0,0,contest.num_question,5+num_usuarios))
  requests.append(add_condition_rule(sheet_id,0,5,contest.num_question,5+num_usuarios))
  for i in range(0,contest.num_question):
    requests.append(add_value(sheet_id,i,4,chr(65+i)))
  body = {"requests":requests}
  spreadsheet.batchUpdate(spreadsheetId=SAMPLE_SPREADSHEET_ID, body=body).execute()


def main():
  creds = None
  creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

  service = build('sheets', 'v4', credentials=creds)
  spreadsheet = service.spreadsheets()

  teste = Contest("caio","caio","caio","caio","caio",3)
  insert_contest(spreadsheet,0,teste)


if __name__ == '__main__':
  main()

