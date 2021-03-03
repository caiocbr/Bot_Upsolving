from googleapiclient.discovery import build
from google.oauth2 import service_account
from google_sheets_utilities import *
from Auth import *
from insert_contest_constants import *

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
    requests.append(request_insert_rows(sheet_id,CONTEST_START_ROW,
                CONTEST_START_ROW+contest.num_question))
    requests.append(request_merge_columns(sheet_id,CONTEST_START_ROW,CONTEST_START_COLUMN,
                CONTEST_START_ROW+contest.num_question,QUESTION_START_ROW))
    requests.append(request_add_hyperlink(sheet_id,CONTEST_START_ROW,CONTEST_START_COLUMN,
                contest.name,contest.link))
    if contest.editorial_link != "":
        requests.append(request_add_hyperlink(sheet_id,CONTEST_START_ROW,CONTEST_START_COLUMN+1,
                    "Editorial",contest.editorial_link))
    requests.append(request_add_value(sheet_id,CONTEST_START_ROW,CONTEST_START_COLUMN+2,
                contest.code))
    requests.append(request_add_value(sheet_id,CONTEST_START_ROW,CONTEST_START_COLUMN+3,
                contest.date))
    for i in range(0,contest.num_question):
        requests.append(request_add_value(sheet_id,QUESTION_START_ROW+i,QUESTION_START_COLUMN,
                    chr(65+i)))
    if num_usuarios > 0:
        formula = get_formula(spreadsheet,sheet_id,PENDING_QUESTION_START_ROW+1,
                PENDING_QUESTION_START_COLUMN)
        aux = ""
        for i in range(0,len(formula)):
            if formula[i] == ':':
                i = i+2
                while formula[i] != ';':
                    aux += formula[i]
                    i = i+1
        for i in range(0,num_usuarios):
            string = '=CONT.SE(' + chr(72+i) + '5:' + chr(72+i) + aux + ';"x")*(-1)+' + str(int(aux)-5)
            requests.append(request_add_formula_value(sheet_id,PENDING_QUESTION_START_ROW,
                        PENDING_QUESTION_START_COLUMN+i,string))
    requests.append(request_centralize_cells(sheet_id,PENDING_QUESTION_START_ROW,
                CONTEST_START_COLUMN,CONTEST_START_ROW+contest.num_question,
                PENDING_QUESTION_START_COLUMN+num_usuarios))
    requests.append(request_border_all_cells(sheet_id,CONTEST_START_ROW,CONTEST_START_COLUMN,
                CONTEST_START_ROW+contest.num_question,
                PENDING_QUESTION_START_COLUMN+num_usuarios))
    requests.append(request_change_color(sheet_id,CONTEST_START_ROW,CONTEST_START_COLUMN,
                CONTEST_START_ROW+contest.num_question,CONTEST_START_ROW+2,0,0,0))
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

