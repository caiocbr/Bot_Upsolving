from googleapiclient.discovery import build
from google.oauth2 import service_account

#transforma um id de coluna (começando por 0) em string de coluna de no maximo 2 caracteres (exemplo: 51->AZ)
def int_to_col(x):
    if(x<26):
        return chr(ord('A')+x)
    else:
        return chr(ord('A')+int(x/26)-1)+chr(ord('A')+x-int(x/26)*26)


#exemplo: mark_question(sheet,SAMPLE_SPREADSHEET_ID,"Deaga","Codeforces Round #701 (Div. 2)","C")
#Se não der certo pode ser porque colocaram um \n no inicio do nome do contest na planilha
#sem querer, vi alguns que estão assim
def mark_question(sheet,SAMPLE_SPREADSHEET_ID,user_name,contest_name,question_name):

    #pegando a coluna do usuario
    request_get = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
     range="p1!H2:ZZ2").execute()
    arr=request_get.get('values',[])
    x=7
    user_col=1000

    for i in arr:
        for j in i:
            if(j==user_name):
                user_col=x
            x=x+1

    if(user_col==1000):
        print("user not found")

    else:

        #pegando as linhas do contest
        request_get = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
         range="p1!A4:A5000").execute()
        arr=request_get.get('values',[])
        x=4
        contest_row_begin=5000
        contest_row_end=5000

        for i in arr:
            for j in i:
                if(contest_row_begin<5000):
                    contest_row_end=x
                    break
                if(j==contest_name):
                    contest_row_begin=x
            if(contest_row_end<5000):
                break
            x=x+1

        if(contest_row_begin==5000):
            print("contest not found")
        
        else:

            #pegando a linha da questão
            request_get = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
             range="p1!G"+str(contest_row_begin)+":G"+str(contest_row_end-1)).execute()
            arr=request_get.get('values',[])
            x=contest_row_begin
            question_row=5000
            for i in arr:
                for j in i:
                    if(j==question_name):
                        question_row=x
                    x=x+1
                    
            if(question_row==5000):
                print("question not found")
                
            else:
                request_upd = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="p1!"+int_to_col(user_col)+str(question_row), 
                 valueInputOption="USER_ENTERED", body={"values":[['X']]}).execute()