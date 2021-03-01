from google_sheets_constants import*

def int_to_column(id):
    ''' Pega uma ID de coluna (começando por 0)
        Retorna a coluna na notação A1 (até ZZ) '''
    if id < 26:
        return chr(ord('A') + id)
    else:
        return chr(ord('A') + int(id / 26) - 1) + chr(ord('A') + id - int(id / 26) * 26)


def get_cells(sheet, sample_spread_sheet_id, range):
    ''' Pega a ID da planilha e o range na notação "p1!A1:ZZ2" 
        Retorna uma matriz das células desse range '''
    request = sheet.values().get(
        spreadsheetId = sample_spread_sheet_id,
        range = range
    ).execute()
    return request.get('values', [])


def get_users_end_column_id(sheet, sample_spread_sheet_id):
    ''' Pega a ID da planilha e retorna a ID da última coluna de usários '''
    range = (
        METADATA_TAB_NAME + "!"
        + int_to_column(USERS_QUANTITY_COLUMN_ID) + str(USERS_QUANTITY_ROW) + ":"
        + int_to_column(USERS_QUANTITY_COLUMN_ID) + str(USERS_QUANTITY_ROW)
    )
    users_quantity = get_cells(sheet, sample_spread_sheet_id, range)
    users_end_column_id = 0
    for i in users_quantity:
        for j in i:
            users_end_column_id = int(j) + USERS_START_COLUMN_ID - 1
    return users_end_column_id


def get_problems_end_row(sheet, sample_spread_sheet_id):
    ''' Pega a ID da planilha e retorna a última linha de problemas '''
    range = (
        METADATA_TAB_NAME + "!"
        + int_to_column(PROBLEMS_QUANTITY_COLUMN_ID) + str(PROBLEMS_QUANTITY_ROW) + ":"
        + int_to_column(PROBLEMS_QUANTITY_COLUMN_ID) + str(PROBLEMS_QUANTITY_ROW)
    )
    problems_quantity = get_cells(sheet, sample_spread_sheet_id, range)
    problems_end_row = 0
    for i in problems_quantity:
        for j in i:
            problems_end_row = int(j) + PROBLEMS_START_ROW - 1
    return problems_end_row


def get_user_column_id(sheet, sample_spread_sheet_id, user_name):
    ''' Pega a ID da planilha e o nome do usuário
        Retorna a ID da coluna do usuário ou None caso não seja encontrado '''
    users_end_column_id = get_users_end_column_id(sheet, sample_spread_sheet_id)
    range = (UPSOLVING_TAB_NAME + "!" 
        + int_to_column(USERS_START_COLUMN_ID) + str(USERS_ROW) + ":" 
        + int_to_column(users_end_column_id) + str(USERS_ROW))
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
    ''' Pega a ID da planilha e o nome da prova
        Retorna a primeira linha da prova ou None caso não seja encontrada '''
    problems_end_row = get_problems_end_row(sheet, sample_spread_sheet_id)
    range = (UPSOLVING_TAB_NAME + "!"
        + CONTESTS_COLUMN + str(PROBLEMS_START_ROW) + ":"
        + CONTESTS_COLUMN + str(problems_end_row))
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
    ''' Pega a ID da planilha e o nome da prova
        Retorna a última linha da prova ou None caso não seja encontrada '''
    problems_end_row = get_problems_end_row(sheet, sample_spread_sheet_id)
    range = (UPSOLVING_TAB_NAME + "!" 
        + CONTESTS_COLUMN + str(PROBLEMS_START_ROW) + ":"
        + CONTESTS_COLUMN + str(problems_end_row))
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
        contest_last_row = problems_end_row

    return contest_last_row


def get_problem_row(sheet, sample_spread_sheet_id, contest_first_row, contest_last_row, problem_name):
    ''' Pega a ID da planilha, a primeira linha da prova, a última linha da prova e o nome da questão
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


def insert_columns(sheet, sample_spread_sheet_id, tab_id , start_column_id , column_quantity):
    ''' Pega a ID da planilha, a ID da aba da planilha,
        a ID da coluna onde começará a inserção e a quantidade de colunas que quer-se inserir
        A ID da aba esta no final do URL, logo depois do 'gid='
        Insere colunas vazias '''
    request = {
        "insertDimension": {
            "range": {
                "sheetId": tab_id ,
                "dimension": "COLUMNS",
                "startIndex": start_column_id,
                "endIndex": start_column_id + column_quantity
            }
        }
    }
    body = {"requests": request}
    sheet.batchUpdate(spreadsheetId=sample_spread_sheet_id, body=body).execute()


def delete_columns(sheet, sample_spread_sheet_id, tab_id , start_column_id , column_quantity):
    ''' Pega a ID da planilha, a ID da aba da planilha,
        a ID da coluna onde começará a deleção e a quantidade de colunas que quer-se deletar
        A ID da aba esta no final do URL, logo depois do 'gid=' '''
    request = {
        "deleteDimension": {
            "range": {
                "sheetId": tab_id ,
                "dimension": "COLUMNS",
                "startIndex": start_column_id,
                "endIndex": start_column_id + column_quantity
            }
        }
    }
    body = {"requests": request}
    sheet.batchUpdate(spreadsheetId=sample_spread_sheet_id, body=body).execute()


def problems_pattern_format(sheet, sample_spread_sheet_id, tab_id, start_row, start_column_id, end_row, end_column_id):
    ''' Pega a ID da planilha, a ID da aba da planilha,
        a linha e a ID da coluna onde começará a formatação 
        e a linha e a ID da coluna onde terminará a formatação no padrão dos problemas
        Vermelho se tiver vazio, Verde se tiver preenchido por X e Branco caso contrário
        A ID da aba esta no final do URL, logo depois do 'gid=' '''
    request = []
    start_row = start_row - 1
    end_column_id = end_column_id + 1

    # Deixa vermelho se a célula estiver vazia
    request_body = {
        "addConditionalFormatRule": {
            "rule": {
                "ranges":{
                    "sheetId": tab_id,
                    "startRowIndex": start_row,
                    "endRowIndex": end_row,
                    "startColumnIndex": start_column_id,
                    "endColumnIndex": end_column_id
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
    request.append(request_body)

    # Deixa verde quando estiver preenchido por X
    request_body = {
        "addConditionalFormatRule": {
            "rule": {
                "ranges":{
                    "sheetId": tab_id,
                    "startRowIndex": start_row,
                    "endRowIndex": end_row,
                    "startColumnIndex": start_column_id,
                    "endColumnIndex": end_column_id
                },
                "booleanRule": {
                    "condition": {
                        "type": "TEXT_CONTAINS",
                        "values": [{"userEnteredValue": "X"}]
                    },
                    "format": {
                        "backgroundColor": {
                            "red": 0.72,
                            "blue": 0.81,
                            "green": 0.88
                        }
                    }
                }
            },
            "index": 0
        }
    }
    request.append(request_body)
    body = {"requests": request}
    sheet.batchUpdate(spreadsheetId=sample_spread_sheet_id, body=body).execute()


def change_column_size(sheet, sample_spread_sheet_id, tab_id, start_column_id, column_quantity, size):
    ''' Pega a ID da planilha, a ID da aba da planilha,
        a ID da coluna onde começará a formatação, a quantidade de colunas a serem formatadas
        e o tamanho desejado para as colunas
        Deixa as colunas selecionadas com o tamanho selecionado
        A ID da aba esta no final do URL, logo depois do 'gid=' '''
    end_column_id = start_column_id + column_quantity
    request = [{
        "updateDimensionProperties": {
            "range": {
                "sheetId": tab_id,
                "dimension": "COLUMNS",
                "startIndex": start_column_id,
                "endIndex": end_column_id
            },
            "properties": {
                "pixelSize": size
            },
            "fields": "pixelSize"
        }
    }]
    body = {"requests": request}
    sheet.batchUpdate(spreadsheetId=sample_spread_sheet_id, body=body).execute()


def centralize_cells_request(tab_id , start_row , start_column_id , end_row , end_column_id):
    ''' Pega a ID da planilha, a ID da aba da planilha,
        a linha e a ID da coluna onde começará a formatação 
        e a linha e a ID da coluna onde terminará a formatação
        Retorna um pedido para centralizar o texto dessas células
        A ID da aba esta no final do URL, logo depois do 'gid=' '''
    start_row = start_row - 1
    end_column_id = end_column_id + 1
    request = {
        "repeatCell":
        {
            "cell":
            {
                "userEnteredFormat":
                {
                    "horizontalAlignment": "CENTER",   
                    "verticalAlignment": "MIDDLE" 
                }
            },
            "range":
            {
                "sheetId": tab_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_column_id,
                "endColumnIndex": end_column_id
            },
            "fields": "userEnteredFormat"
        }
    }
    return request


def centralize_cells(sheet, sample_spread_sheet_id, tab_id , start_row , start_column_id , end_row , end_column_id):
    ''' Pega a ID da planilha, a ID da aba da planilha,
        a linha e a ID da coluna onde começará a formatação 
        e a linha e a ID da coluna onde terminará a formatação
        O texto dessas células será centralizado
        A ID da aba esta no final do URL, logo depois do 'gid=' '''
    request = [centralize_cells_request(tab_id, start_row, start_column_id, end_row, end_column_id)]
    body = {"requests": request}
    sheet.batchUpdate(spreadsheetId=sample_spread_sheet_id, body=body).execute()


def bold_cells_request(tab_id, start_row, start_column_id, end_row, end_column_id):
    ''' Pega a ID da planilha, a ID da aba da planilha,
        a linha e a ID da coluna onde começará a formatação 
        e a linha e a ID da coluna onde terminará a formatação
        Retorna um pedido para colocar em negrito o texto dessas células
        A ID da aba esta no final do URL, logo depois do 'gid=' '''
    start_row = start_row - 1
    end_column_id = end_column_id + 1
    request = {
        "repeatCell": {
            "range":{
                "sheetId": tab_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_column_id,
                "endColumnIndex": end_column_id
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
    return request


def bold_cells(sheet, sample_spread_sheet_id, tab_id, start_row, start_column_id, end_row, end_column_id):
    ''' Pega a ID da planilha, a ID da aba da planilha,
        a linha e a ID da coluna onde começará a formatação 
        e a linha e a ID da coluna onde terminará a formatação
        O texto dessas células será colocado em negrito
        A ID da aba esta no final do URL, logo depois do 'gid=' '''
    request = [bold_cells_request(tab_id, start_row, start_column_id, end_row, end_column_id)]
    body = {"requests": request}
    sheet.batchUpdate(spreadsheetId=sample_spread_sheet_id, body=body).execute()


def external_border_request(tab_id, start_row, start_column_id, end_row, end_column_id):
    ''' Pega a ID da aba da planilha,
        a linha e a ID da coluna onde começará a formatação 
        e a linha e a ID da coluna onde terminará a formatação
        Retorna um pedido para tornar a borda externa do grupo de células escolhido realçada
        A ID da aba esta no final do URL, logo depois do 'gid=' '''
    start_row = start_row - 1
    end_column_id = end_column_id + 1
    request = {
        "updateBorders": {
            "range":{
                "sheetId": tab_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_column_id,
                "endColumnIndex": end_column_id
            },
            "top": {"style":"SOLID"},
            "bottom": {"style":"SOLID"},
            "left": {"style":"SOLID"},
            "right": {"style":"SOLID"}
        }
    }
    return request


def external_border(sheet, sample_spread_sheet_id, tab_id, start_row, start_column_id, end_row, end_column_id):
    ''' Pega a ID da planilha, a ID da aba da planilha,
        a linha e a ID da coluna onde começará a formatação 
        e a linha e a ID da coluna onde terminará a formatação
        A borda externa do grupo de células escolhido será realçada
        A ID da aba esta no final do URL, logo depois do 'gid=' '''
    start_row = start_row - 1
    end_column_id = end_column_id + 1
    request = [external_border_request(tab_id,start_row, start_column_id, end_row, end_column_id)]
    body = {"requests": request}
    sheet.batchUpdate(spreadsheetId=sample_spread_sheet_id, body=body).execute()


def set_text_rotation_request(tab_id, start_row, start_column_id, end_row, end_column_id, angle):
    ''' Pega a ID da planilha, a ID da aba da planilha,
        a linha e a ID da coluna onde começará a formatação 
        e a linha e a ID da coluna onde terminará a formatação
        e quantos graus se deseja rotacionar o Texto das células selecionadas
        Retorna um pedido para rotacionar o Texto das células selecionadas nos graus selecionados
        A ID da aba esta no final do URL, logo depois do 'gid=' '''
    start_row = start_row - 1
    end_column_id = end_column_id + 1
    request = {
        "repeatCell": {
            "range":{
                "sheetId":  tab_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_column_id,
                "endColumnIndex": end_column_id
            },
            "cell": {
                "userEnteredFormat": {
                    "textRotation": {
                        "angle" : angle
                    }
                }
            },
            "fields": "userEnteredFormat"
        }
    }
    return request


def set_text_rotation(sheet, sample_spread_sheet_id, tab_id, start_row, start_column_id, end_row, end_column_id, angle):
    ''' Pega a ID da planilha, a ID da aba da planilha,
        a linha e a ID da coluna onde começará a formatação 
        e a linha e a ID da coluna onde terminará a formatação
        e quantos graus se deseja rotacionar o Texto das células selecionadas
        Rotaciona o Texto das células selecionadas nos graus selecionados
        A ID da aba esta no final do URL, logo depois do 'gid=' '''
    request = [set_text_rotation_request(tab_id, start_row, start_column_id, end_row, end_column_id, angle)]
    body = {"requests": request}
    sheet.batchUpdate(spreadsheetId=sample_spread_sheet_id, body=body).execute()


def write_request(tab_id, row, column_id, text):
    ''' Pega a ID da aba da planilha, a linha e a ID da coluna onde se quer escrever
        e o texto que se quer colocar nesta célula
        Retorna um pedido para escrever o texto na célula escolhida
        A ID da aba esta no final do URL, logo depois do 'gid=' '''
    row = row - 1
    request = {
        "updateCells": 
        {
            "rows": 
            {
                "values": {
                    "userEnteredValue": {
                        "stringValue": text
                    }
                }
            },
            "fields": "userEnteredValue",
            "start": {
                "sheetId": tab_id,
                "rowIndex": row,
                "columnIndex": column_id 
            }
        }
    }
    return request


def user_pattern_format(sheet, sample_spread_sheet_id, tab_id, row, column_id):
    ''' Pega a ID da aba da planilha, a linha e a ID da coluna que se quer formatar
        Deixa no formato dos usuários, deitado em 45º e centralizado
        A ID da aba esta no final do URL, logo depois do 'gid=' '''
    request = [{
        "repeatCell": {
            "range":{
                "sheetId":  tab_id,
                "startRowIndex": row - 1,
                "endRowIndex": row,
                "startColumnIndex": column_id,
                "endColumnIndex": column_id + 1
            },
            "cell": {
                "userEnteredFormat": {
                    "horizontalAlignment": "CENTER",   
                    "verticalAlignment": "MIDDLE",
                    "textRotation": {
                        "angle" : 45
                    }
                }
            },
            "fields": "userEnteredFormat"
        }
    }]
    body = {"requests": request}
    sheet.batchUpdate(spreadsheetId=sample_spread_sheet_id, body=body).execute()