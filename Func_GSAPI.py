# Todas as funcoes apenas compoem um request, apos usa-las e necessarios executar
# a seguinte linha de codigo para que as alteracoes sejam aplicadas:
# body = {"requests":requests}
# spreadsheet.batchUpdate(spreadsheetId=SAMPLE_SPREADSHEET_ID, body=body).execute()

# Insere uma "num" linhas na linha de numero "start_index"
def insert_rows(sheet_id , start_index , num):
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

# Junta as linhas
# num grid com o canto superior esquerdo em (pos1,pos2) e canto inferior direito em (pos3,pos4)
def merge_columns(sheet_id , pos1 , pos2 , pos3 , pos4):
  request_body = {
    "mergeCells": {
      "range": {
        "sheetId": sheet_id,
        "startRowIndex": pos1,
        "endRowIndex": pos3,
        "startColumnIndex": pos2,
        "endColumnIndex": pos4
      },
      "mergeType": "MERGE_COLUMNS"
    }
  }
  return request_body

# Junta as colunas
# num grid com o canto superior esquerdo em (pos1,pos2) e canto inferior direito em (pos3,pos4)
def merge_rows(sheet_id , pos1 , pos2 , pos3 , pos4):
  request_body = {
    "mergeCells": {
      "range": {
        "sheetId": sheet_id,
        "startRowIndex": pos1,
        "endRowIndex": pos3,
        "startColumnIndex": pos2,
        "endColumnIndex": pos4
      },
      "mergeType": "MERGE_ROWS"
    }
  }
  return request_body

# Centraliza o texto 
# num grid com o canto superior esquerdo em (pos1,pos2) e canto inferior direito em (pos3,pos4)
def centralize_cells(sheet_id , pos1 , pos2 , pos3 , pos4):
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
        "startRowIndex": pos1,
        "endRowIndex": pos3,
        "startColumnIndex": pos2,
        "endColumnIndex": pos4
      },
      "fields": "userEnteredFormat"
    }
  }
  return request_body

# Adiciona um texto(name) e uma url(url) na celula (pos1,pos2)
def add_hyperlink(sheet_id , pos1 , pos2 , name , url):
  request_body = {
    "updateCells": 
    {
      "rows": 
        [
        {
          "values": [{
            "userEnteredValue": {
              "formulaValue":'=HYPERLINK("{}";"{}")'.format(url,name)
            }
          }]
        }
        ],
        "fields": "userEnteredValue",
        "start": {
          "sheetId": sheet_id,
          "rowIndex": pos1,
          "columnIndex": pos2 
        }
    }
  }
  return request_body

# Escreve "value" na celula (pos1,pos2)
def add_value(sheet_id , pos1 , pos2 , value):
  request_body = {
    "updateCells": 
    {
      "rows": 
        [
        {
          "values": [{
            "userEnteredValue": {
              "stringValue":value
            }
          }]
        }
        ],
        "fields": "userEnteredValue",
        "start": {
          "sheetId": sheet_id,
          "rowIndex": pos1,
          "columnIndex": pos2 
        }
    }
  }
  return request_body

# Realca apenas as bordas externas
# num grid com o canto superior esquerdo em (pos1,pos2) e canto inferior direito em (pos3,pos4)
def border_cell(sheet_id , pos1 , pos2 , pos3 , pos4):
  request_body = {
    "updateBorders": {
      "range":{
        "sheetId": sheet_id,
        "startRowIndex": pos1,
        "endRowIndex": pos3,
        "startColumnIndex": pos2,
        "endColumnIndex": pos4
      },
      "top": {"style":"SOLID"},
      "bottom": {"style":"SOLID"},
      "left": {"style":"SOLID"},
      "right": {"style":"SOLID"}
    }
  }
  return request_body

# Realca todas as bordas
# num grid com o canto superior esquerdo em (pos1,pos2) e canto inferior direito em (pos3,pos4)
def border_all_cells(sheet_id , pos1 , pos2 , pos3 , pos4):
  request_body = {
    "repeatCell": {
      "range":{
        "sheetId": sheet_id,
        "startRowIndex": pos1,
        "endRowIndex": pos3,
        "startColumnIndex": pos2,
        "endColumnIndex": pos4
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

# Muda a cor do texto
# num grid com o canto superior esquerdo em (pos1,pos2) e canto inferior direito em (pos3,pos4)
# para o sistema rgb,com cada parametro variando de [0,1]
def change_color(sheet_id , pos1 , pos2 , pos3 , pos4 , red , green , blue):
  request_body = {
    "repeatCell": {
      "range":{
        "sheetId": sheet_id,
        "startRowIndex": pos1,
        "endRowIndex": pos3,
        "startColumnIndex": pos2,
        "endColumnIndex": pos4
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

# Add a regra de deixar vermelho se estiver em branco
# num grid com o canto superior esquerdo em (pos1,pos2) e canto inferior direito em (pos3,pos4)
def add_condition_rule1(sheet_id , pos1 , pos2 , pos3 , pos4):
  request_body = {
  "addConditionalFormatRule": {
     "rule": {
       "ranges":{
         "sheetId": sheet_id,
         "startRowIndex": pos1,
         "endRowIndex": pos3,
         "startColumnIndex": pos2,
         "endColumnIndex": pos4
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

# Add a regra de deixar verde se nao estiver vazio
# num grid com o canto superior esquerdo em (pos1,pos2) e canto inferior direito em (pos3,pos4)
def add_condition_rule2(sheet_id , pos1 , pos2 , pos3 , pos4):
  request_body = {
    "addConditionalFormatRule": {
     "rule": {
       "ranges":{
         "sheetId": sheet_id,
         "startRowIndex": pos1,
        "endRowIndex": pos3,
        "startColumnIndex": pos2,
        "endColumnIndex": pos4
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

# Altera o tamanho das colunas comecando na posicao "startIndex" e temrinando na posicao
# endIndex para size
def change_size_column(sheet_id , startIndex , endIndex , size):
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

# Altera o tamanho das linhas comecando na posicao "startIndex" e temrinando na posicao
# endIndex para size
def change_size_row(sheet_id , startIndex , endIndex , size):
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

# Altera o texto para ficar em negrito
# num grid com o canto superior esquerdo em (pos1,pos2) e canto inferior direito em (pos3,pos4)
def bold_cells(sheet_id , pos1 , pos2 , pos3 , pos4):
  request_body = {
    "repeatCell": {
      "range":{
        "sheetId": sheet_id,
        "startRowIndex": pos1,
        "endRowIndex": pos3,
        "startColumnIndex": pos2,
        "endColumnIndex": pos4
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

