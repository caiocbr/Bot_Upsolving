from google_sheets_constants import*

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