# Todas as funcoes apenas compoem um request, apos usa-las e necessarios executar
# a seguinte linha de codigo para que as alteracoes sejam aplicadas:
# requests.append(func)
# body = {"requests":requests}
# spreadsheet.batchUpdate(spreadsheetId=SAMPLE_SPREADSHEET_ID, body=body).execute()

def insert_rows(sheet_id , start_index , num):
    '''Insere uma "num" linhas na linha de numero "start_index"'''
    request_body = {
        "insertDimension": {
            "range": {
                "sheetId": sheet_id ,
                "dimension": "ROWS",
                "startIndex": start_index,
                "endIndex": num
            }
        }
    }
    return request_body

def merge_columns(sheet_id , startRowIndex , startColumnIndex , endRowIndex , endColumnIndex):
    '''Junta as linhas num grid com o canto superior esquerdo em 
       (startRow,startColumn) e canto inferior direito em (endRow,endColumn)'''
    request_body = {
        "mergeCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": startRowIndex,
                "endRowIndex": endRowIndex,
                "startColumnIndex": startColumnIndex,
                "endColumnIndex": endColumnIndex
            },
            "mergeType": "MERGE_COLUMNS"
        }
    }
    return request_body

def merge_rows(sheet_id , startRowIndex , startColumnIndex , endRowIndex , endColumnIndex):
    '''Junta as colunas num grid com o canto superior esquerdo em 
       (startRow,startColumn) e canto inferior direito em (endRow,endColumn)'''
    request_body = {
        "mergeCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": startRowIndex,
                "endRowIndex": endRowIndex,
                "startColumnIndex": startColumnIndex,
                "endColumnIndex": endColumnIndex
            },
            "mergeType": "MERGE_ROWS"
        }
    }
    return request_body

def centralize_cells(sheet_id , startRowIndex , startColumnIndex , endRowIndex , endColumnIndex):
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
                "startRowIndex": startRowIndex,
                "endRowIndex": endRowIndex,
                "startColumnIndex": startColumnIndex,
                "endColumnIndex": endColumnIndex
            },
            "fields": "userEnteredFormat"
        }
    }
    return request_body

def add_hyperlink(sheet_id , rowIndex , columnIndex , name , url):
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
                    "rowIndex": rowIndex,
                    "columnIndex": columnIndex
                }
        }
    }
    return request_body

def add_value(sheet_id , rowIndex , columnIndex , value):
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
                "rowIndex": rowIndex,
                "columnIndex": columnIndex 
            }
        }
    }
    return request_body

def border_cell(sheet_id , startRowIndex , startColumnIndex , endRowIndex , endColumnIndex):
    '''Realca apenas as bordas externas num grid com o canto superior esquerdo em 
       (startRow,startColumn) e canto inferior direito em (endRow,endColumn)'''
    request_body = {
        "updateBorders": {
            "range":{
                "sheetId": sheet_id,
                "startRowIndex": startRowIndex,
                "endRowIndex": endRowIndex,
                "startColumnIndex": startColumnIndex,
                "endColumnIndex": endColumnIndex
            },
            "top": {"style":"SOLID"},
            "bottom": {"style":"SOLID"},
            "left": {"style":"SOLID"},
            "right": {"style":"SOLID"}
        }
    }
    return request_body

def border_all_cells(sheet_id , startRowIndex , startColumnIndex , endRowIndex , endColumnIndex):
    '''Realca todas as bordas num grid com o canto superior esquerdo em 
       (startRow,startColumn) e canto inferior direito em (endRow,endColumn)'''
    request_body = {
        "repeatCell": {
            "range":{
                "sheetId": sheet_id,
                "startRowIndex": startRowIndex,
                "endRowIndex": endRowIndex,
                "startColumnIndex": startColumnIndex,
                "endColumnIndex": endColumnIndex
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

def change_color(sheet_id , startRowIndex , startColumnIndex , endRowIndex , endColumnIndex , red , green , blue):
    '''Muda a cor do texto num grid com o canto superior esquerdo em 
       (startRow,startColumn) e canto inferior direito em (endRow,endColumn)
       para o sistema rgb,com cada parametro variando de [0,1]'''
    request_body = {
        "repeatCell": {
            "range":{
                "sheetId": sheet_id,
                "startRowIndex": startRowIndex,
                "endRowIndex": endRowIndex,
                "startColumnIndex": startColumnIndex,
                "endColumnIndex": endColumnIndex
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

def add_condition_rule1(sheet_id , startRowIndex , startColumnIndex , endRowIndex , endColumnIndex):
    '''Add a regra de deixar vermelho se estiver em branco num grid com o canto superior 
       esquerdo em (startRow,startColumn) e canto inferior direito em (endRow,endColumn)'''
    request_body = {
        "addConditionalFormatRule": {
            "rule": {
                "ranges":{
                    "sheetId": sheet_id,
                    "startRowIndex": startRowIndex,
                    "endRowIndex": endRowIndex,
                    "startColumnIndex": startColumnIndex,
                    "endColumnIndex": endColumnIndex
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

def add_condition_rule2(sheet_id , startRowIndex , startColumnIndex , endRowIndex , endColumnIndex):
    '''Add a regra de deixar verde se nao estiver vazio num grid com o canto superior 
       esquerdo em (startRow,startColumn) e canto inferior direito em (endRow,endColumn)'''
    request_body = {
        "addConditionalFormatRule": {
            "rule": {
                "ranges":{
                    "sheetId": sheet_id,
                    "startRowIndex": startRowIndex,
                    "endRowIndex": endRowIndex,
                    "startColumnIndex": startColumnIndex,
                    "endColumnIndex": endColumnIndex
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

def change_size_column(sheet_id , startIndex , endIndex , size):
    '''Altera o tamanho das colunas comecando na posicao "startIndex" e temrinando na posicao
       endIndex para size'''
    request_body = {
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "COLUMNS",
                "startIndex": startIndex,
                "endIndex": endIndex
            },
            "properties": {
                "pixelSize": size
            },
            "fields": "pixelSize"
        }
    }
    return request_body

def change_size_row(sheet_id , startIndex , endIndex , size):
    '''Altera o tamanho das linhas comecando na posicao "startIndex" e temrinando na posicao 
       endIndex para size'''
    request_body = {
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "ROWS",
                "startIndex": startIndex,
                "endIndex": endIndex
            },
            "properties": {
                "pixelSize": size
            },
            "fields": "pixelSize"
        }
    }
    return request_body

def bold_cells(sheet_id , startRowIndex , startColumnIndex , endRowIndex , endColumnIndex):
    '''Altera o texto para ficar em negrito num grid com o canto superior esquerdo em 
       (startRow,startColumn) e canto inferior direito em (endRow,endColumn)'''
    request_body = {
        "repeatCell": {
            "range":{
                "sheetId": sheet_id,
                "startRowIndex": startRowIndex,
                "endRowIndex": endRowIndex,
                "startColumnIndex": startColumnIndex,
                "endColumnIndex": endColumnIndex
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

def add_formula_value(sheet_id , rowIndex , columnIndex , value):
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
                "rowIndex": rowIndex,
                "columnIndex": columnIndex 
            }
        }
    }
    return request_body
