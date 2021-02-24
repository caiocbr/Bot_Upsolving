from google_sheets_constants import*

# ID de coluna (começando por 0) -> notação A1 (até ZZ)
def int_to_column(id):
    if id < 26:
        return chr(ord('A') + id)
    else:
        return chr(ord('A') + int(id / 26) - 1) + chr(ord('A') + id - int(id / 26) * 26)


# Range na notação "p1!A1:ZZ2"
def get_cells(sheet, SAMPLE_SPREADSHEET_ID, range):
    request = sheet.values().get(
        spreadsheetId = SAMPLE_SPREADSHEET_ID,
        range = range
    ).execute()
    return request.get('values',[])


def get_user_column(sheet, SAMPLE_SPREADSHEET_ID, user_name):
    users = get_cells(sheet,
        SAMPLE_SPREADSHEET_ID, 
        Upsolving_tab_name + "!" 
        + int_to_column(Users_start_column_id) + str(Users_row) + ":" 
        + int_to_column(Users_end_column_id) + str(Users_row)
    )
    column = Users_start_column_id
    user_column = None

    for i in users:
        for j in i:
            if j == user_name:
                user_column = column
            column = column + 1
    
    return user_column


def get_contest_first_row(sheet, SAMPLE_SPREADSHEET_ID, contest_name):
    contests = get_cells(sheet, 
        SAMPLE_SPREADSHEET_ID, 
        Upsolving_tab_name + "!"
        + Contests_column + str(Problems_start_row) + ":"
        + Contests_column + str(Problems_end_row)
    )
    row = Problems_start_row
    contest_first_row = None
    
    for i in contests:
        for j in i:
            if j == contest_name:
                contest_first_row = row
        row = row + 1

    return contest_first_row


def get_contest_last_row(sheet, SAMPLE_SPREADSHEET_ID, contest_name):
    contests = get_cells(sheet, 
        SAMPLE_SPREADSHEET_ID,
        Upsolving_tab_name + "!" 
        + Contests_column + str(Problems_start_row) + ":"
        + Contests_column + str(Problems_end_row)
    )
    row = Problems_start_row
    aux = False
    contest_last_row = None
    
    for i in contests:
        for j in i:
            if j == contest_name:
                aux = True
            elif aux:
                contest_last_row = row
                break
        if contest_last_row != None:
            break
        row = row + 1

    return contest_last_row


def get_problem_row(sheet, SAMPLE_SPREADSHEET_ID, contest_first_row, contest_last_row, problem_name):
    problems = get_cells(sheet, 
        SAMPLE_SPREADSHEET_ID, 
        Upsolving_tab_name + "!" 
        + Problems_column + str(contest_first_row) + ":" 
        + Problems_column + str(contest_last_row-1)
    )
    row = contest_first_row
    problem_row = None

    for i in problems:
        for j in i:
            if j == problem_name:
                problem_row = row
        row = row + 1
    
    return problem_row