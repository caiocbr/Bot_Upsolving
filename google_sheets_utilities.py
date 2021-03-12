from google_sheets_constants import *
from Auth import *

def insert_rows_request(tab_id, start_index, rows_quantity):
    ''' Pega o ID da tabela, a linha de inicio e a quantidade de linhas a serem inseridas
        Retorna o request que insere uma "num" linhas na linha de numero "start_index"'''
    request_body = {
        "insertDimension": {
            "range": {
                "sheetId": tab_id ,
                "dimension": "ROWS",
                "startIndex": start_index,
                "endIndex": start_index+rows_quantity
            }
        }
    }
    return request_body


def merge_columns_request(tab_id, start_row, start_column, end_row, end_column):
    ''' Pega o ID da tabela e o range de aplicacao da funcao
        Retorna o request que junta as linhas num grid com o canto superior esquerdo em 
        (start_row, start_column) e canto inferior direito em (end_row-1, end_column-1)'''
    request_body = {
        "mergeCells": {
            "range": {
                "sheetId": tab_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_column,
                "endColumnIndex": end_column
            },
            "mergeType": "MERGE_COLUMNS"
        }
    }
    return request_body


def merge_rows_request(tab_id, start_row, start_column, end_row, end_column):
    ''' Pega o ID da tabela e o range de aplicacao da funcao
        Retorna o request que junta as colunas num grid com o canto superior esquerdo em 
        (start_row, start_column) e canto inferior direito em (end_row-1, end_column-1)'''
    request_body = {
        "mergeCells": {
            "range": {
                "sheetId": tab_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_column,
                "endColumnIndex": end_column
            },
            "mergeType": "MERGE_ROWS"
        }
    }
    return request_body


def centralize_cells_request(tab_id, start_row, start_column, end_row, end_column):
    ''' Pega o ID da tabela e o range de aplicacao da funcao
        Retorna o request que centraliza o texto num grid com o canto superior esquerdo em 
        (start_row, start_column) e canto inferior direito em (end_row-1, end_column-1)'''
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
                "sheetId": tab_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_column,
                "endColumnIndex": end_column
            },
            "fields": "userEnteredFormat"
        }
    }
    return request_body


def add_hyperlink_request(tab_id, row_index, column_index, name, url):
    ''' Pega o ID da tabela, a posicao da celula e o nome e a url a ser inserido 
        Retorna o request que adiciona um texto(name) e uma url(url) na 
        celula (row_index, column_index)'''
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
                    "sheetId": tab_id,
                    "rowIndex": row_index,
                    "columnIndex": column_index
                }
        }
    }
    return request_body


def add_value_request(tab_id, row_index, column_index, value):
    ''' Pega o ID da tabela, a posicao da celula e o valor a ser escrito 
        Retorna o request que escreve "value" na celula (row_index, column_index)'''
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
                "sheetId": tab_id,
                "rowIndex": row_index,
                "columnIndex": column_index 
            }
        }
    }
    return request_body


def border_cell_request(tab_id, start_row, start_column, end_row, end_column):
    ''' Pega o ID da tabela e o range de aplicacao da funcao
        Retorna o request que realca apenas as bordas externas num grid com o canto superior 
        esquerdo em (start_row, start_column) e canto inferior direito em 
        (end_row-1, end_column-1)'''
    request_body = {
        "updateBorders": {
            "range":{
                "sheetId": tab_id,
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


def border_all_cells_request(tab_id, start_row, start_column, end_row, end_column):
    ''' Pega o ID da tabela e o range de aplicacao da funcao
        Retorna o request que realca todas as bordas num grid com o canto superior esquerdo em 
        (start_row, start_column) e canto inferior direito em (end_row-1, end_column-1)'''
    request_body = {
        "repeatCell": {
            "range":{
                "sheetId": tab_id,
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


def change_color_request(tab_id, start_row, start_column, end_row, end_column, red, green, blue):
    ''' Pega o ID da tabela, o range de aplicacao da funcao e as cores RGB
        Retorna o request que muda a cor do texto num grid com o canto superior esquerdo em 
        (start_row, start_column) e canto inferior direito em (end_row-1, end_column-1)
        para o sistema rgb,com cada parametro variando de [0,1]'''
    request_body = {
        "repeatCell": {
            "range":{
                "sheetId": tab_id,
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


def add_condition_rule_red_blank_request(tab_id, start_row, start_column, end_row, end_column):
    ''' Pega o ID da tabela e o range de aplicacao da funcao
        Retorna o request que add a regra de deixar vermelho se estiver em branco 
        num grid com o canto superior esquerdo em (start_row, start_column) e 
        canto inferior direito em (end_row-1, end_column-1)'''
    request_body = {
        "addConditionalFormatRule": {
            "rule": {
                "ranges":{
                    "sheetId": tab_id,
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


def add_condition_rule_green_not_blank_request(tab_id, start_row, start_column, end_row, end_column):
    ''' Pega o ID da tabela e o range de aplicacao da funcao
        Retorna o request que add a regra de deixar verde se nao estiver vazio num grid com 
        o canto superior esquerdo em (start_row, start_column) e canto inferior direito 
        em (end_row-1, end_column-1)'''
    request_body = {
        "addConditionalFormatRule": {
            "rule": {
                "ranges":{
                    "sheetId": tab_id,
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


def change_size_column_request(tab_id, start_index, end_index, size):
    ''' Pega o ID da tabela, a linha de inicio e de fim e o tamanho da coluna desejado 
        Retorna o request que altera o tamanho das colunas comecando na posicao 
        "start_index" e temrinando na posicao end_index para size'''
    request_body = {
        "updateDimensionProperties": {
            "range": {
                "sheetId": tab_id,
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


def change_size_row_request(tab_id, start_index, end_index, size):
    ''' Pega o ID da tabela, a coluna de inicio e de fim e o tamanho da linha desjado
        Retorna o request que altera o tamanho das linhas comecando na posicao "start_index" 
        e temrinando na posicao end_index para size'''
    request_body = {
        "updateDimensionProperties": {
            "range": {
                "sheetId": tab_id,
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


def bold_cells_request(tab_id, start_row, start_column, end_row, end_column):
    ''' Pega o ID da tabela e o range de aplicacao da funcao
        Retorna o request que altera o texto para ficar em negrito num grid com o 
        canto superior esquerdo em (start_row, start_column) e canto inferior direito 
        em (end_row-1, end_column-1)'''
    request_body = {
        "repeatCell": {
            "range":{
                "sheetId": tab_id,
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


def add_formula_value_request(tab_id, row_index, column_index, value):
    ''' Pega o ID da tabela, a posicao da celula e o valor a ser escrito
        Retorna o request que adiciona uma formula na celula de posicao 
        (row_index, column_index) com o valor "value"'''
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
                "sheetId": tab_id,
                "rowIndex": row_index,
                "columnIndex": column_index 
            }
        }
    }
    return request_body


def get_formula(spreadsheet, name_tab, row_index, column_index):
    ''' Pega a spreadsheet, o nome da tabela e a posicao da celula
        Retorna a formula da celula de posicao (row_index,column_index)'''
    result = spreadsheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range= name_tab + 
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
