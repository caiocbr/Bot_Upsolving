from google_sheets_utilities import*
from google_sheets_constants import*

def change_problem_cell(sheet, spreadsheet_id, user_name, contest_name, problem_name, mark):
    ''' Pega a ID da planilha, o nome do usuário, o nome da prova, o nome do problema e como se deseja marcar
        Retorna uma string avisando que marcou a questão do jeito desejado no local desejado
        ou avisa que o usuário, a prova ou a questão não foram encontrados '''
    
    # Pegando a coluna do usuário
    user_column = get_user_column_id(sheet, spreadsheet_id, user_name)
    if user_column == None:
        return "Usuário não encontrado"

    # Pegando as linhas da prova
    contest_first_row = get_contest_first_row(sheet, spreadsheet_id, contest_name)
    contest_last_row = get_contest_last_row(sheet, spreadsheet_id, contest_name)
    if contest_first_row == None:
        return "Prova não encontrada"

    # Pegando a linha do problema
    problem_row = get_problem_row(sheet, 
        spreadsheet_id, 
        contest_first_row, 
        contest_last_row,
        problem_name
    )        
    if problem_row == None:
        return "Problema não encontrado"
            
    sheet.values().update(
        spreadsheetId = spreadsheet_id, 
        range = UPSOLVING_TAB_NAME + "!" + int_to_column(user_column) + str(problem_row), 
        valueInputOption = "USER_ENTERED", 
        body = {"values":[[mark]]}
    ).execute()
    if mark == 'X':
        return "Questão marcada!"
    elif mark == 'N':
        return "Questão marcada como 'por enquanto não'!"
    else:
        return "Questão desmarcada!"


def mark_problem(sheet, spreadsheet_id, user_name, contest_name, problem_name):
    ''' Pega a ID da planilha, o nome do usuário, o nome da prova e o nome do problema
        Retorna uma string avisando que marcou a questão feita no local desejado
        ou avisa que o usuário, a prova ou a questão não foram encontrados '''

    return change_problem_cell(sheet, spreadsheet_id, user_name, contest_name, problem_name, 'X')


def desmark_problem(sheet, spreadsheet_id, user_name, contest_name, problem_name):
    ''' Pega a ID da planilha, o nome do usuário, o nome da prova e o nome do problema
        Retorna uma string avisando que desmarcou a questão no local desejado
        ou avisa que o usuário, a prova ou a questão não foram encontrados '''

    return change_problem_cell(sheet, spreadsheet_id, user_name, contest_name, problem_name, '')


def not_for_now_problem(sheet, spreadsheet_id, user_name, contest_name, problem_name):
    ''' Pega a ID da planilha, o nome do usuário, o nome da prova e o nome do problema
        Retorna uma string avisando que marcou a questão com N no local desejado
        ou avisa que o usuário, a prova ou a questão não foram encontrados '''

    return change_problem_cell(sheet, spreadsheet_id, user_name, contest_name, problem_name, 'N')


def add_user(sheet, spreadsheet_id, tab_id, user_name):
    ''' Pega a ID da planilha e o nome do novo usuário
        Retorna uma string avisando que adcionou o usuário no início da aba de Upsolving 
        ou avisando que já tem um usuário com esse nome '''

    user_exist = get_user_column_id(sheet, spreadsheet_id, user_name)
    if(user_exist != None):
        return "Já existe um usuário com esse nome"
    
    problems_end_row = get_problems_end_row(sheet, spreadsheet_id)

    # Pegando a quantidade de usuários
    range = (
        METADATA_TAB_NAME + "!"
        + int_to_column(USERS_QUANTITY_COLUMN_ID) + str(USERS_QUANTITY_ROW) + ":"
        + int_to_column(USERS_QUANTITY_COLUMN_ID) + str(USERS_QUANTITY_ROW)
    )
    users_quantity = 0
    cells = get_cells(sheet, spreadsheet_id, range)
    for i in cells:
        for j in i:
            users_quantity = int(j)
       
    # Inserindo coluna no início
    insert_columns(sheet, spreadsheet_id, tab_id, USERS_START_COLUMN_ID, 1)

    # Formatando a coluna caso seja o primeiro usuário a ser adicionado
    if users_quantity == 0:
        format_problem_cells(
            sheet, spreadsheet_id, tab_id,
            PROBLEMS_START_ROW, USERS_START_COLUMN_ID,
            problems_end_row, USERS_START_COLUMN_ID
        )
        change_column_size(sheet, spreadsheet_id, tab_id, USERS_START_COLUMN_ID, 1, USERS_COLUMN_SIZE)
        format_user_cells(sheet, spreadsheet_id, tab_id, USERS_ROW, USERS_START_COLUMN_ID)
        centralize_cells(sheet, spreadsheet_id, tab_id, USERS_ROW + 1, USERS_START_COLUMN_ID, problems_end_row, USERS_START_COLUMN_ID)
        bold_cells(sheet, spreadsheet_id, tab_id, USERS_ROW, USERS_START_COLUMN_ID, PROBLEM_COUNTER_ROW, USERS_START_COLUMN_ID)

        # Colocando as bordas
        range = (
            UPSOLVING_TAB_NAME + "!"
            + CONTESTS_COLUMN + str(PROBLEMS_START_ROW + 1) + ":"
            + CONTESTS_COLUMN + str(problems_end_row)
        )
        cells = get_cells(sheet, spreadsheet_id, range)
        start_row = PROBLEMS_START_ROW
        end_row = PROBLEMS_START_ROW
        requests = []
        for i in cells:
            for j in i:
                requests.append(external_border_request(tab_id, start_row, USERS_START_COLUMN_ID, end_row, USERS_START_COLUMN_ID))
                start_row = end_row + 1
            end_row = end_row + 1
        if problems_end_row >= PROBLEMS_START_ROW:
            requests.append(external_border_request(tab_id, start_row, USERS_START_COLUMN_ID, problems_end_row, USERS_START_COLUMN_ID))
        requests.append(external_border_request(tab_id, USERS_ROW, USERS_START_COLUMN_ID, USERS_ROW, USERS_START_COLUMN_ID))
        body = {"requests": requests}
        sheet.batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

    # Inserindo o nome do usuário
    sheet.values().update(
        spreadsheetId = spreadsheet_id, 
        range = UPSOLVING_TAB_NAME + "!" + int_to_column(USERS_START_COLUMN_ID) + str(USERS_ROW), 
        valueInputOption = "USER_ENTERED", 
        body = {"values":[[user_name]]}
    ).execute()

    # Inserindo lógica de contar questão
    logica = ("=SE(ÉCÉL.VAZIA(" + int_to_column(USERS_START_COLUMN_ID) + str(USERS_ROW) 
        + "); ""; CONTAR.VAZIO(" + int_to_column(USERS_START_COLUMN_ID) + str(PROBLEMS_START_ROW) + ":" 
        + int_to_column(USERS_START_COLUMN_ID) + str(problems_end_row) +"))")
    sheet.values().update(
        spreadsheetId = spreadsheet_id, 
        range = UPSOLVING_TAB_NAME + "!" + int_to_column(USERS_START_COLUMN_ID) + str(PROBLEM_COUNTER_ROW), 
        valueInputOption = "USER_ENTERED", 
        body = {"values":[[logica]]}
    ).execute()

    # Colocando N em todas as questões que existiam antes de adicionar este usuário
    row = PROBLEMS_START_ROW
    requests = []
    while row <= problems_end_row:
        requests.append(write_request(tab_id,row,USERS_START_COLUMN_ID,"N"))
        row = row + 1
    body = {"requests": requests}
    sheet.batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

    # Atualizando a quantidade de usuários
    new_users_quantity = users_quantity + 1
    sheet.values().update(
        spreadsheetId = spreadsheet_id, 
        range = METADATA_TAB_NAME + "!" + int_to_column(USERS_QUANTITY_COLUMN_ID) + str(USERS_QUANTITY_ROW), 
        valueInputOption = "USER_ENTERED", 
        body = {"values":[[new_users_quantity]]}
    ).execute()
    return "Adicionado!"


def delete_user(sheet, spreadsheet_id, tab_id, user_name):
    ''' Pega a ID da planilha e o nome do usuário que quer deletar
        Retorna uma string avisando que deletou a coluna do usuário 
        ou avisa que o usuário não foi encontrado '''

    #deletando a coluna do usuário
    user_column_id = get_user_column_id(sheet, spreadsheet_id, user_name)
    if(user_column_id == None):
        return "Usuário não encontrado"
    
    delete_columns(sheet, spreadsheet_id, tab_id, user_column_id, 1)
        
    # Atualizando a quantidade de usuários
    range = (
        METADATA_TAB_NAME + "!"
        + int_to_column(USERS_QUANTITY_COLUMN_ID) + str(USERS_QUANTITY_ROW) + ":"
        + int_to_column(USERS_QUANTITY_COLUMN_ID) + str(USERS_QUANTITY_ROW)
    )
    users_quantity = get_cells(sheet, spreadsheet_id, range)
    new_users_quantity = 0
    for i in users_quantity:
        for j in i:
            new_users_quantity = int(j) - 1
    sheet.values().update(
        spreadsheetId = spreadsheet_id, 
        range = METADATA_TAB_NAME + "!" + int_to_column(USERS_QUANTITY_COLUMN_ID) + str(USERS_QUANTITY_ROW), 
        valueInputOption = "USER_ENTERED", 
        body = {"values":[[new_users_quantity]]}
    ).execute()
    return "Deletado!"