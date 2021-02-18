def insert_rows(sheet_id , num):
  request_body = {
    "insertDimension": {
      "range": {
        "sheetId": sheet_id ,
        "dimension": "ROWS",
        "startIndex": 0,
        "endIndex": num
      }
    }
  }
  return request_body

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

#Apenas as bordas externas
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

def add_condition_rule(sheet_id , pos1 , pos2 , pos3 , pos4):
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
     }
   },
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
     }
   }
  }
  return request_body
