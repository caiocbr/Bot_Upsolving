from google_sheets_utilities import*

def mark_problem(sheet, SAMPLE_SPREADSHEET_ID, user_name, contest_name, problem_name):
    user_column = get_user_column_id(sheet, SAMPLE_SPREADSHEET_ID, user_name)
    if user_column == None:
        print("user not found")

    else:
        contest_first_row = get_contest_first_row(sheet, SAMPLE_SPREADSHEET_ID, contest_name)
        contest_last_row = get_contest_last_row(sheet, SAMPLE_SPREADSHEET_ID, contest_name)
        if contest_first_row == None:
            print("contest not found")

        else:
            problem_row = get_problem_row(sheet, 
                SAMPLE_SPREADSHEET_ID, 
                contest_first_row, 
                contest_last_row,
                problem_name
            )        
            if problem_row == None:
                print("problem not found")
                
            else:
                request_upd = sheet.values().update(
                    spreadsheetId = SAMPLE_SPREADSHEET_ID, 
                    range = "p1!" + int_to_column(user_column) + str(problem_row), 
                    valueInputOption = "USER_ENTERED", 
                    body = {"values":[['X']]}
                ).execute()