from google_sheets_utilities import*
from google_sheets_constants import*

def mark_problem(sheet, spread_sheet_id, user_name, contest_name, problem_name):
    user_column = get_user_column_id(sheet, spread_sheet_id, user_name)
    if user_column == None:
        print("user not found")

    else:
        contest_first_row = get_contest_first_row(sheet, spread_sheet_id, contest_name)
        contest_last_row = get_contest_last_row(sheet, spread_sheet_id, contest_name)
        if contest_first_row == None:
            print("contest not found")

        else:
            problem_row = get_problem_row(sheet, 
                spread_sheet_id, 
                contest_first_row, 
                contest_last_row,
                problem_name
            )        
            if problem_row == None:
                print("problem not found")
                
            else:
                sheet.values().update(
                    spreadsheetId = spread_sheet_id, 
                    range = UPSOLVING_TAB_NAME + "!" + int_to_column(user_column) + str(problem_row), 
                    valueInputOption = "USER_ENTERED", 
                    body = {"values":[['X']]}
                ).execute()