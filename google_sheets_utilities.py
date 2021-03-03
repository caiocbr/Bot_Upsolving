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
