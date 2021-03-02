from google_sheets_utilities import*
from google_sheets_constants import*

def add_user(sheet, spread_sheet_id, user_name):
    ''' Pega a ID da planilha e o nome do novo usuário
        Adiciona usuário no início da aba de Upsolving ou avisa que já tem um usuário com esse nome '''
    user_exist = get_user_column_id(sheet, spread_sheet_id, user_name)
    if(user_exist != None):
        print("Já existe um usuário com esse nome")
    else:
        problems_end_row = get_problems_end_row(sheet, spread_sheet_id)
        # Pegando a quantidade de usuários
        range = (
            METADATA_TAB_NAME + "!"
            + int_to_column(USERS_QUANTITY_COLUMN_ID) + str(USERS_QUANTITY_ROW) + ":"
            + int_to_column(USERS_QUANTITY_COLUMN_ID) + str(USERS_QUANTITY_ROW)
        )
        users_quantity = 0
        cells = get_cells(sheet, spread_sheet_id, range)
        for i in cells:
            for j in i:
                users_quantity = int(j)
        
        # Inserindo coluna no início
        insert_columns(sheet, spread_sheet_id, UPSOLVING_TAB_ID, USERS_START_COLUMN_ID, 1)

        # Formatando a coluna caso seja o primeiro usuário a ser adicionado
        if users_quantity == 0:
            format_problem_cells(
                sheet, spread_sheet_id, UPSOLVING_TAB_ID,
                PROBLEMS_START_ROW, USERS_START_COLUMN_ID,
                problems_end_row, USERS_START_COLUMN_ID
            )
            change_column_size(sheet, spread_sheet_id, UPSOLVING_TAB_ID, USERS_START_COLUMN_ID, 1, USERS_COLUMN_SIZE)
            format_user_cells(sheet, spread_sheet_id, UPSOLVING_TAB_ID, USERS_ROW, USERS_START_COLUMN_ID)
            centralize_cells(sheet, spread_sheet_id, UPSOLVING_TAB_ID, USERS_ROW + 1, USERS_START_COLUMN_ID, problems_end_row, USERS_START_COLUMN_ID)
            bold_cells(sheet, spread_sheet_id, UPSOLVING_TAB_ID, USERS_ROW, USERS_START_COLUMN_ID, PROBLEM_COUNTER_ROW, USERS_START_COLUMN_ID)

            # Colocando as bordas
            range = (
                UPSOLVING_TAB_NAME + "!"
                + CONTESTS_COLUMN + str(PROBLEMS_START_ROW + 1) + ":"
                + CONTESTS_COLUMN + str(problems_end_row)
            )
            cells = get_cells(sheet, spread_sheet_id, range)
            start_row = PROBLEMS_START_ROW
            end_row = PROBLEMS_START_ROW
            requests = []
            for i in cells:
                for j in i:
                    requests.append(external_border_request(UPSOLVING_TAB_ID, start_row, USERS_START_COLUMN_ID, end_row, USERS_START_COLUMN_ID))
                    start_row = end_row + 1
                end_row = end_row + 1
            if problems_end_row >= PROBLEMS_START_ROW:
                requests.append(external_border_request(UPSOLVING_TAB_ID, start_row, USERS_START_COLUMN_ID, problems_end_row, USERS_START_COLUMN_ID))
            requests.append(external_border_request(UPSOLVING_TAB_ID, USERS_ROW, USERS_START_COLUMN_ID, USERS_ROW, USERS_START_COLUMN_ID))
            body = {"requests": requests}
            sheet.batchUpdate(spreadsheetId=spread_sheet_id, body=body).execute()

        # Inserindo o nome do usuário
        sheet.values().update(
            spreadsheetId = spread_sheet_id, 
            range = UPSOLVING_TAB_NAME + "!" + int_to_column(USERS_START_COLUMN_ID) + str(USERS_ROW), 
            valueInputOption = "USER_ENTERED", 
            body = {"values":[[user_name]]}
        ).execute()

        # Inserindo lógica de contar questão
        logica = ("=SE(ÉCÉL.VAZIA(" + int_to_column(USERS_START_COLUMN_ID) + str(USERS_ROW) 
            + "); ""; CONTAR.VAZIO(" + int_to_column(USERS_START_COLUMN_ID) + str(PROBLEMS_START_ROW) + ":" 
            + int_to_column(USERS_START_COLUMN_ID) + str(problems_end_row) +"))")
        sheet.values().update(
            spreadsheetId = spread_sheet_id, 
            range = UPSOLVING_TAB_NAME + "!" + int_to_column(USERS_START_COLUMN_ID) + str(PROBLEM_COUNTER_ROW), 
            valueInputOption = "USER_ENTERED", 
            body = {"values":[[logica]]}
        ).execute()

        # Colocando N em todas as questões que existiam antes de adicionar este usuário
        row = PROBLEMS_START_ROW
        requests = []
        while row <= problems_end_row:
            requests.append(write_request(UPSOLVING_TAB_ID,row,USERS_START_COLUMN_ID,"N"))
            row = row + 1
        body = {"requests": requests}
        sheet.batchUpdate(spreadsheetId=spread_sheet_id, body=body).execute()

        # Atualizando a quantidade de usuários
        new_users_quantity = users_quantity + 1
        sheet.values().update(
            spreadsheetId = spread_sheet_id, 
            range = METADATA_TAB_NAME + "!" + int_to_column(USERS_QUANTITY_COLUMN_ID) + str(USERS_QUANTITY_ROW), 
            valueInputOption = "USER_ENTERED", 
            body = {"values":[[new_users_quantity]]}
        ).execute()

def delete_user(sheet, spread_sheet_id, user_name):
    ''' Pega a ID da planilha e o nome do usuário que quer deletar
        Deleta a coluna do Usuário ou avisa que o usuário não foi encontrado '''
    #deletando a coluna do usuário
    user_column_id = get_user_column_id(sheet, spread_sheet_id, user_name)
    if(user_column_id == None):
        print("Usuário não encontrado")
    else:
        delete_columns(sheet, spread_sheet_id, UPSOLVING_TAB_ID, user_column_id, 1)
        
        # Atualizando a quantidade de usuários
        range = (
            METADATA_TAB_NAME + "!"
            + int_to_column(USERS_QUANTITY_COLUMN_ID) + str(USERS_QUANTITY_ROW) + ":"
            + int_to_column(USERS_QUANTITY_COLUMN_ID) + str(USERS_QUANTITY_ROW)
        )
        users_quantity = get_cells(sheet, spread_sheet_id, range)
        new_users_quantity = 0
        for i in users_quantity:
            for j in i:
                new_users_quantity = int(j) - 1
        sheet.values().update(
            spreadsheetId = spread_sheet_id, 
            range = METADATA_TAB_NAME + "!" + int_to_column(USERS_QUANTITY_COLUMN_ID) + str(USERS_QUANTITY_ROW), 
            valueInputOption = "USER_ENTERED", 
            body = {"values":[[new_users_quantity]]}
        ).execute()