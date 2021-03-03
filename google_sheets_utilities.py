from google_sheets_constants import *
from Auth import *

def request_insert_rows(sheet_id, start_index, num_rows):
    '''Insere uma "num" linhas na linha de numero "start_index"'''
    request_body = {
        "insertDimension": {
            "range": {
                "sheetId": sheet_id ,
                "dimension": "ROWS",
                "startIndex": start_index,
                "endIndex": num_rows
            }
        }
    }
    return request_body


def request_merge_columns(sheet_id, start_row, start_column, end_row, end_column):
    '''Junta as linhas num grid com o canto superior esquerdo em 
       (startRow,startColumn) e canto inferior direito em (endRow,endColumn)'''
    request_body = {
        "mergeCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_column,
                "endColumnIndex": end_column
            },
            "mergeType": "MERGE_COLUMNS"
        }
    }
    return request_body


def request_merge_rows(sheet_id, start_row, start_column, end_row, end_column):
    '''Junta as colunas num grid com o canto superior esquerdo em 
       (startRow,startColumn) e canto inferior direito em (endRow,endColumn)'''
    request_body = {
        "mergeCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_column,
                "endColumnIndex": end_column
            },
            "mergeType": "MERGE_ROWS"
        }
    }
    return request_body


def request_centralize_cells(sheet_id, start_row, start_column, end_row, end_column):
    '''Centraliza o texto num grid com o canto superior esquerdo em 
       (startRow,startColumn) e canto inferior direito em (endRow,endColumn)'''
    request_body = {
        "repeatCell":
        {
            "cell":
            {
                "userEnteredFormat":
                {
                    "horizontalAlignment": "CENTER" ,   
                    "verticalAlignment": "MIDDLE" 
                }
            },
            "range":
            {
                "sheetId": sheet_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_column,
                "endColumnIndex": end_column
            },
            "fields": "userEnteredFormat"
        }
    }
    return request_body


def request_add_hyperlink(sheet_id, row_index, column_index, name, url):
    '''Adiciona um texto(name) e uma url(url) na celula (rowIndex,columnIndex)'''
    request_body = {
        "updateCells": 
        {
            "rows": 
            {
                "values": {
                    "userEnteredValue": {
                        "formulaValue":'=HYPERLINK("{}";"{}")'.format(url,name)
                    }
                }
            }
            ,
                "fields": "userEnteredValue",
                "start": {
                    "sheetId": sheet_id,
                    "rowIndex": row_index,
                    "columnIndex": column_index
                }
        }
    }
    return request_body


def request_add_value(sheet_id, row_index, column_index, value):
    '''Escreve "value" na celula (rowIndex,columnIndex)'''
    request_body = {
        "updateCells": 
        {
            "rows": 
            {
                "values": {
                    "userEnteredValue": {
                        "stringValue": value
                    }
                }
            },
            "fields": "userEnteredValue",
            "start": {
                "sheetId": sheet_id,
                "rowIndex": row_index,
                "columnIndex": column_index 
            }
        }
    }
    return request_body


def request_border_cell(sheet_id, start_row, start_column, end_row, end_column):
    '''Realca apenas as bordas externas num grid com o canto superior esquerdo em 
       (startRow,startColumn) e canto inferior direito em (endRow,endColumn)'''
    request_body = {
        "updateBorders": {
            "range":{
                "sheetId": sheet_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_column,
                "endColumnIndex": end_column
            },
            "top": {"style":"SOLID"},
            "bottom": {"style":"SOLID"},
            "left": {"style":"SOLID"},
            "right": {"style":"SOLID"}
        }
    }
    return request_body


def request_border_all_cells(sheet_id, start_row, start_column, end_row, end_column):
    '''Realca todas as bordas num grid com o canto superior esquerdo em 
       (startRow,startColumn) e canto inferior direito em (endRow,endColumn)'''
    request_body = {
        "repeatCell": {
            "range":{
                "sheetId": sheet_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_column,
                "endColumnIndex": end_column
            },
            "cell": {
                "userEnteredFormat": {
                    "borders": {
                        "top": {"style":"SOLID"},
                        "bottom": {"style":"SOLID"},
                        "left": {"style":"SOLID"},
                        "right": {"style":"SOLID"}
                    }
                }
            },
            "fields": "userEnteredFormat.borders"
        }
    }
    return request_body


def request_change_color(sheet_id, start_row, start_column, end_row, end_column, red, green, blue):
    '''Muda a cor do texto num grid com o canto superior esquerdo em 
       (startRow,startColumn) e canto inferior direito em (endRow,endColumn)
       para o sistema rgb,com cada parametro variando de [0,1]'''
    request_body = {
        "repeatCell": {
            "range":{
                "sheetId": sheet_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_column,
                "endColumnIndex": end_column
            },
            "cell": {
                "userEnteredFormat": {
                    "textFormat": {
                        "foregroundColor": {
                            "red": red,
                            "green": green,
                            "blue": blue
                        }
                    }
                }
            },
            "fields": "userEnteredFormat.textFormat"
        }
    }
    return request_body


def request_add_condition_rule_red_blank(sheet_id, start_row, start_column, end_row, end_column):
    '''Add a regra de deixar vermelho se estiver em branco num grid com o canto superior 
       esquerdo em (startRow,startColumn) e canto inferior direito em (endRow,endColumn)'''
    request_body = {
        "addConditionalFormatRule": {
            "rule": {
                "ranges":{
                    "sheetId": sheet_id,
                    "startRowIndex": start_row,
                    "endRowIndex": end_row,
                    "startColumnIndex": start_column,
                    "endColumnIndex": end_column
                },
                "booleanRule": {
                    "condition": {"type": "BLANK"},
                    "format": {
                        "backgroundColor": {
                            "red": 1, 
                            "blue": 0.8,
                            "green": 0.8
                        }
                    }
                }
            },
            "index": 0
        }
    }
    return request_body


def request_add_condition_rule_green_not_blank(sheet_id, start_row, start_column, end_row, end_column):
    '''Add a regra de deixar verde se nao estiver vazio num grid com o canto superior 
       esquerdo em (startRow,startColumn) e canto inferior direito em (endRow,endColumn)'''
    request_body = {
        "addConditionalFormatRule": {
            "rule": {
                "ranges":{
                    "sheetId": sheet_id,
                    "startRowIndex": start_row,
                    "endRowIndex": end_row,
                    "startColumnIndex": start_column,
                    "endColumnIndex": end_column
                },
                "booleanRule": {
                    "condition": {"type": "NOT_BLANK"},
                    "format": {
                        "backgroundColor": {
                            "red": 0.8,
                            "blue": 0.8,
                            "green": 1
                        }
                    }
                }
            },
            "index": 0
        }
    }
    return request_body


def request_change_size_column(sheet_id, start_index, end_index, size):
    '''Altera o tamanho das colunas comecando na posicao "startIndex" e temrinando na posicao
       endIndex para size'''
    request_body = {
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "COLUMNS",
                "startIndex": start_index,
                "endIndex": end_index
            },
            "properties": {
                "pixelSize": size
            },
            "fields": "pixelSize"
        }
    }
    return request_body


def request_change_size_row(sheet_id, start_index, end_index, size):
    '''Altera o tamanho das linhas comecando na posicao "startIndex" e temrinando na posicao 
       endIndex para size'''
    request_body = {
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "ROWS",
                "startIndex": start_index,
                "endIndex": end_index
            },
            "properties": {
                "pixelSize": size
            },
            "fields": "pixelSize"
        }
    }
    return request_body


def request_bold_cells(sheet_id, start_row, start_column, end_row, end_column):
    '''Altera o texto para ficar em negrito num grid com o canto superior esquerdo em 
       (startRow,startColumn) e canto inferior direito em (endRow,endColumn)'''
    request_body = {
        "repeatCell": {
            "range":{
                "sheetId": sheet_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_column,
                "endColumnIndex": end_column
            },
            "cell": {
                "userEnteredFormat": {
                    "textFormat": {
                        "bold" : True
                    }
                }
            },
            "fields": "userEnteredFormat.textFormat"
        }
    }
    return request_body


def request_add_formula_value(sheet_id, row_index, column_index, value):
    '''Adiciona uma formula na celula de posicao (rowIndex,columnIndex) com o valor "value"'''
    request_body = {
        "updateCells": 
        {
            "rows": 
            {
                "values": {
                    "userEnteredValue" : {
                        "formulaValue" : value
                    }
                }
            },
            "fields": "userEnteredValue",
            "start": {
                "sheetId": sheet_id,
                "rowIndex": row_index,
                "columnIndex": column_index 
            }
        }
    }
    return request_body


def get_formula(spreadsheet, sheet_id, row_index, column_index):
    result = spreadsheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range= "p1!" + 
            chr(65+column_index) + str(row_index) + ":"+ chr(65+column_index) + 
            str(row_index),valueRenderOption = "FORMULA").execute()
    values = result.get('values', [])
    return values[0][0]


def int_to_column(id):
    ''' Pega uma ID de coluna (começando por 0)
        Retorna a coluna na notação A1 (até ZZ) '''
    if id < 26:
        return chr(ord('A') + id)
    else:
        return chr(ord('A') + int(id / 26) - 1) + chr(ord('A') + id - int(id / 26) * 26)


def get_cells(sheet, sample_spread_sheet_id, range):
    ''' Pega a ID da tabela e o range na notação "p1!A1:ZZ2" 
        Retorna uma matriz das células desse range '''
    request = sheet.values().get(
        spreadsheetId = sample_spread_sheet_id,
        range = range
    ).execute()
    return request.get('values', [])


def get_user_column_id(sheet, sample_spread_sheet_id, user_name):
    ''' Pega a ID da tabela e o nome do usuário
        Retorna a ID da coluna do usuário ou None caso não seja encontrado '''
    range = (UPSOLVING_TAB_NAME + "!" 
        + int_to_column(USERS_START_COLUMN_ID) + str(USERS_ROW) + ":" 
        + int_to_column(USERS_END_COLUMN_ID) + str(USERS_ROW))
    users = get_cells(sheet,
        sample_spread_sheet_id, 
        range
    )
    column = USERS_START_COLUMN_ID
    user_column = None

    for i in users:
        for j in i:
            if j == user_name:
                user_column = column
            column = column + 1
    
    return user_column


def get_contest_first_row(sheet, sample_spread_sheet_id, contest_name):
    ''' Pega a ID da tabela e o nome da prova
        Retorna a primeira linha da prova ou None caso não seja encontrada '''
    range = (UPSOLVING_TAB_NAME + "!"
        + CONTESTS_COLUMN + str(PROBLEMS_START_ROW) + ":"
        + CONTESTS_COLUMN + str(PROBLEMS_END_ROW))
    contests = get_cells(sheet, 
        sample_spread_sheet_id, 
        range
    )
    row = PROBLEMS_START_ROW
    contest_first_row = None
    
    for i in contests:
        for j in i:
            if j == contest_name:
                contest_first_row = row
        row = row + 1

    return contest_first_row


def get_contest_last_row(sheet, sample_spread_sheet_id, contest_name):
    ''' Pega a ID da tabela e o nome da prova
        Retorna a última linha da prova ou None caso não seja encontrada '''
    range = (UPSOLVING_TAB_NAME + "!" 
        + CONTESTS_COLUMN + str(PROBLEMS_START_ROW) + ":"
        + CONTESTS_COLUMN + str(PROBLEMS_END_ROW))
    contests = get_cells(sheet, 
        sample_spread_sheet_id,
        range
    )
    row = PROBLEMS_START_ROW
    aux = False
    contest_last_row = None
    
    for i in contests:
        for j in i:
            if j == contest_name:
                aux = True
            elif aux:
                contest_last_row = row - 1
                break
        if contest_last_row != None:
            break
        row = row + 1
    
    if aux and (contest_last_row == None):
        contest_last_row = PROBLEMS_END_ROW

    return contest_last_row


def get_problem_row(sheet, sample_spread_sheet_id, contest_first_row, contest_last_row, problem_name):
    ''' Pega a ID da tabela, a primeira linha da prova, a última linha da prova e o nome da questão
        Retorna a linha da questão ou None caso não seja encontrada '''
    range = (UPSOLVING_TAB_NAME + "!" 
        + PROBLEMS_COLUMN + str(contest_first_row) + ":" 
        + PROBLEMS_COLUMN + str(contest_last_row))
    problems = get_cells(sheet, 
        sample_spread_sheet_id, 
        range
    )
    row = contest_first_row
    problem_row = None

    for i in problems:
        for j in i:
            if j == problem_name:
                problem_row = row
        row = row + 1
    
    return problem_row
