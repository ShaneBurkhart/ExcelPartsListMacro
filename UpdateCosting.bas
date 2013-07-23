'Basic Macro to update all of the costs on the master list
'Creator: Shane Burkhart

'Constants
Public Const MAX_ROWS As Long = 65536
Public Const NOT_PROJECT_PROMPT As String = "You are not on a project page."
Public Const NOT_PROJECT_TITLE As String = "Not A Valid Page"
Public Const NO_PART_NUM_PROMPT As String = "Part number does not exist: "
Public Const NO_PART_NUM_TITLE As String = "Part Number Doesn't Exist"
Public Const WRONG_MEASURE_PROMPT As String = "Unit of measure does not match for part: "
Public Const WRONG_MEASURE_TITLE As String = "Part Unit Of Measure Mismatch"

Public Const PROJECT_NAME_CELL As String = "B1"
Public Const PROJECT_PART_NUM_COLUMN = "B"
Public Const PROJECT_HAND_COLUMN = "D"
Public Const PROJECT_BUILDING_COLUMN = "F"
Public Const PROJECT_MEASURE_COLUMN = "I"
Public Const PROJECT_UNIT_COLUMN = "G"
Public Const PROJECT_MULTIPLYER_COLUMN = "H" 'Num of takeoff per unit
Public Const PROJECT_DATA_BEGIN = 6

Public Const MASTER_DATA_BEGIN As Integer = 5
Public Const MASTER_PROJECT_COLUMN As String = "A"
Public Const MASTER_PART_NUM_COLUMN As String = "C"
Public Const MASTER_HAND_COLUMN As String = "E"
Public Const MASTER_QUANTITY_COLUMN As String = "G"
Public Const MASTER_BUILDING_COLUMN As String = "J"
Public Const MASTER_FLOOR_COLUMN As String = "K"
Public Const MASTER_DIVISION_COLUMN As String = "B"
Public Const MASTER_MEASURE_COLUMN As String = "H"
Public Const MASTER_UNIT_COST_COLUMN As String = "M"
Public Const MASTER_TOTAL_COST_COLUMN As String = "N"
Public Const MASTER_FLOOR_COST_PSF_COLUMN As String = "O"
Public Const MASTER_BLDG_COST_PSF_COLUMN As String = "P"
Public Const MASTER_SHEET_NAME As String = "Master Parts List"

Public Const VALID_SHEET_NAME As String = "Validation Source Lists"
Public Const VALID_DATA_BEGIN As Integer = 5
Public Const VALID_PROJECT_COLUMN As String = "A"
Public Const VALID_DIVISION_COLUMN As String = "B"

Public Const UNIT_SORT_BEGIN As String = "6"
Public Const UNIT_BASEMENT_STD_COLUMN = "L"
Public Const UNIT_BASEMENT_REV_COLUMN = "M"
Public Const UNIT_FIRST_STD_COLUMN = "O"
Public Const UNIT_FIRST_REV_COLUMN = "P"
Public Const UNIT_SECOND_STD_COLUMN = "R"
Public Const UNIT_SECOND_REV_COLUMN = "S"
Public Const UNIT_THIRD_STD_COLUMN = "U"
Public Const UNIT_THIRD_REV_COLUMN = "V"
Public Const UNIT_FOURTH_STD_COLUMN = "X"
Public Const UNIT_FOURTH_REV_COLUMN = "Y"
Public Const UNIT_GENERAL_STD_COLUMN = "AA"
Public Const UNIT_GENERAL_REV_COLUMN = "AB"

Public Const UNIT_BASEMENT_SF_COLUMN As String = "L"
Public Const UNIT_FIRST_SF_COLUMN As String = "O"
Public Const UNIT_SECOND_SF_COLUMN As String = "R"
Public Const UNIT_THIRD_SF_COLUMN As String = "U"
Public Const UNIT_FOURTH_SF_COLUMN As String = "X"
Public Const UNIT_TOTAL_SF_COLUMN As String = "AA"
Public Const UNIT_SF_ROW As Integer = 1

Public Const PART_NUM_SHEET_NAME = "Part No."
Public Const PART_NUM_DATA_BEGIN = 5
Public Const PART_NUM_PART_NUM_COLUMN = "A"
Public Const PART_NUM_MEASURE_COLUMN = "B"
Public Const PART_NUM_COST_COLUMN = "C"

Public Const HAND_RIGHT = "R"
Public Const HAND_LEFT = "L"


Sub UpdateCosting()

Sheets(MASTER_SHEET_NAME).Unprotect

InsertCosting

Sheets(MASTER_SHEET_NAME).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True

End Sub

Function InsertCosting()
    Dim masterRow As Integer: masterRow = MASTER_DATA_BEGIN
    Dim master_measure As String: Dim part_num As String: Dim part_sheet_row As Integer: Dim arb As Integer
    While Not Sheets(MASTER_SHEET_NAME).Cells(masterRow, MASTER_PROJECT_COLUMN).Value = ""
        master_measure = Sheets(MASTER_SHEET_NAME).Cells(masterRow, MASTER_MEASURE_COLUMN).Value
        part_num = Sheets(MASTER_SHEET_NAME).Cells(masterRow, MASTER_PART_NUM_COLUMN).Value
        part_sheet_row = GetRowOnPartNumSheet(part_num)
        If part_sheet_row = 0 Then
            invalidMessage = MsgBox(NO_PART_NUM_PROMPT + part_num, vbOKOnly, NO_PART_NUM_TITLE)
        Else
            If Sheets(PART_NUM_SHEET_NAME).Cells(part_sheet_row, PART_NUM_MEASURE_COLUMN).Value = master_measure Then
                arb = CostToMaster(part_sheet_row, masterRow)
            Else
               invalidMessage = MsgBox(WRONG_MEASURE_PROMPT + part_num, vbOKOnly, WRONG_MEASURE_TITLE)
            End If
        End If
        masterRow = masterRow + 1
    Wend
End Function

Function CostToMaster(part_sheet_row As Integer, master_row As Integer)
    If Sheets(PART_NUM_SHEET_NAME).Cells(part_sheet_row, PART_NUM_COST_COLUMN).Value = "" Then
        Sheets(MASTER_SHEET_NAME).Cells(master_row, MASTER_UNIT_COST_COLUMN).Value = "NO COST"
    Else
        Dim unit_cost As Double: unit_cost = Sheets(PART_NUM_SHEET_NAME).Cells(part_sheet_row, PART_NUM_COST_COLUMN).Value
        Dim qty As Double: qty = Sheets(MASTER_SHEET_NAME).Cells(master_row, MASTER_QUANTITY_COLUMN).Value
        Dim floor_sf As Long: floor_sf = GetFloorSquareFoot(master_row)
        Dim total_sf As Long: total_sf = GetTotalSquareFoot(master_row)
        Dim total_cost As Double: total_cost = unit_cost * qty
        Sheets(MASTER_SHEET_NAME).Cells(master_row, MASTER_UNIT_COST_COLUMN).Value = unit_cost
        Sheets(MASTER_SHEET_NAME).Cells(master_row, MASTER_TOTAL_COST_COLUMN).Value = total_cost
        Sheets(MASTER_SHEET_NAME).Cells(master_row, MASTER_FLOOR_COST_PSF_COLUMN).Value = total_cost / floor_sf
        Sheets(MASTER_SHEET_NAME).Cells(master_row, MASTER_BLDG_COST_PSF_COLUMN).Value = total_cost / total_sf
    End If
End Function

Function GetFloorSquareFoot(master_row As Integer)
    Dim project As String: project = Sheets(MASTER_SHEET_NAME).Cells(master_row, MASTER_PROJECT_COLUMN).Value
    Dim floor As String: floor = Sheets(MASTER_SHEET_NAME).Cells(master_row, MASTER_FLOOR_COLUMN).Value
    If floor = "B" Then
        GetFloorSquareFoot = Sheets(project).Cells(UNIT_SF_ROW, UNIT_BASEMENT_SF_COLUMN).Value
        Exit Function
    ElseIf floor = "1" Then
        GetFloorSquareFoot = Sheets(project).Cells(UNIT_SF_ROW, UNIT_FIRST_SF_COLUMN).Value
        Exit Function
    ElseIf floor = "2" Then
        GetFloorSquareFoot = Sheets(project).Cells(UNIT_SF_ROW, UNIT_SECOND_SF_COLUMN).Value
        Exit Function
    ElseIf floor = "3" Then
        GetFloorSquareFoot = Sheets(project).Cells(UNIT_SF_ROW, UNIT_THIRD_SF_COLUMN).Value
        Exit Function
    ElseIf floor = "4" Then
        GetFloorSquareFoot = Sheets(project).Cells(UNIT_SF_ROW, UNIT_FOURTH_SF_COLUMN).Value
        Exit Function
    Else
        GetFloorSquareFoot = Sheets(project).Cells(UNIT_SF_ROW, UNIT_TOTAL_SF_COLUMN).Value
        Exit Function
    End If
End Function

Function GetTotalSquareFoot(master_row As Integer)
    Dim project As String: project = Sheets(MASTER_SHEET_NAME).Cells(master_row, MASTER_PROJECT_COLUMN).Value
    GetTotalSquareFoot = Sheets(project).Cells(UNIT_SF_ROW, UNIT_TOTAL_SF_COLUMN).Value
End Function

Function GetRowOnPartNumSheet(part_num As String)
    Dim row As Integer: row = PART_NUM_DATA_BEGIN
    While Not Sheets(PART_NUM_SHEET_NAME).Cells(row, PART_NUM_PART_NUM_COLUMN).Value = ""
        If Sheets(PART_NUM_SHEET_NAME).Cells(row, PART_NUM_PART_NUM_COLUMN).Value = part_num Then
            GetRowOnPartNumSheet = row
            Exit Function
        End If
        row = row + 1
    Wend
    GetRowOnPartNumSheet = 0
End Function
