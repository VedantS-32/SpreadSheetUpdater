import gspread as gs
from gspread.cell import Cell
from gspread.exceptions import APIError
from oauth2client.service_account import ServiceAccountCredentials
from google.oauth2.service_account import Credentials
from enum import Enum
import os

# Google Form Order
# Note while I was testing Columns in Form Spread Sheet
# was zeroth index for some reason
# be aware of this while setting enums
class FormField(Enum):
    TimeStamp = 0
    Name = 1
    Question1 = 2
    Question2 = 3
    Question3 = 4

# Google Sheet Order
# Columns in main worksheet are indexed from 1
class MainField(Enum):
    Name = 1
    QuesA = 2
    QuesB = 3
    QuesC = 4
    LastUpdate = 5

# Checks if Credentials Exists
# if(os.path.exists("sheet_key.json")):
#     print("Exist!\n")

# Credentials
scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# Enter your path to Key.json here
credentials = Credentials.from_service_account_file(
    'KeyFormat.json',
    scopes=scopes
)

# Authorizes Your Credentials
gc = gs.authorize(credentials)

# Opens Google Sheet by Url
sh = gc.open_by_url("")

# Opens Google Form Sheet(Index 0) and Main Sheet(Index 1)
FormWorkSheet = sh.get_worksheet(0)
MainWorkSheet = sh.get_worksheet(1)

# Currently Using Name for UpdateID
UpdateID = FormWorkSheet.col_values(2)

CellToUpdate = []

for i in range(1, len(UpdateID)):
    print(f"------- On {i} / {(len(UpdateID) - 1)} -------")
    UpdateIDCell = FormWorkSheet.findall(UpdateID[i])

    cells = MainWorkSheet.findall(UpdateID[i])
    print("-->Searching", UpdateID[i])
    if len(cells) > 0:
        print('Found', UpdateID[i], 'in Main WorkSheet')
        MainWorkSheetCell = MainWorkSheet.find(UpdateIDCell[-1].value)
    else:
        print(UpdateID[i], "Doesn't Exist Skipping")
        continue

    FormRowValue = FormWorkSheet.row_values(UpdateIDCell[-1].row)
    MainRowValue = MainWorkSheet.row_values(MainWorkSheetCell.row)

    while (len(FormRowValue) <= len(MainField)) :
        FormRowValue.append('')
    print(len(MainRowValue))
    while (len(MainRowValue) <= len(MainField)) :
        MainRowValue.append('')
    
    print("Preparing Update Cells")

    # Update QuestionA Field
    if (FormRowValue[FormField.Question1.value] != ''):
        CellQuestionA = Cell(MainWorkSheetCell.row, MainField.QuesA.value, FormRowValue[FormField.Question1.value])
        CellToUpdate.append(CellQuestionA)

    # Update QuestionB Field
    if (FormRowValue[FormField.Question2.value] != ''):
        CellQuestionB = Cell(MainWorkSheetCell.row, MainField.QuesB.value, FormRowValue[FormField.Question2.value])
        CellToUpdate.append(CellQuestionB)

    # Update QuestionA Field
    if (FormRowValue[FormField.Question3.value] != ''):
        CellQuestionC = Cell(MainWorkSheetCell.row, MainField.QuesC.value, FormRowValue[FormField.Question3.value])
        CellToUpdate.append(CellQuestionC)
    
    #Last Updated
    CellLastUpdate = Cell(MainWorkSheetCell.row, MainField.LastUpdate.value, FormRowValue[FormField.TimeStamp.value])
    CellToUpdate.append(CellLastUpdate)

    MainWorkSheet.update_cells(CellToUpdate)
    print(f"*****Updated Successfully!***** {round((i/(len(UpdateID) - 1)) * 100, 2)}% done")

print("========SpreadSheet Update Completed========")