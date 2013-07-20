'Basic Macro to create Master Parts list from a unit parts list
'Creator: Shane Burkhart

'Left are the check for whether or not there is a hand
'If there is a hand then Left and Right vars will be set appropriately
'However if there is no hand Left will be left as 0 to skip duplicates

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


Sub CreateMasterPartsList()
    'Variables
    Dim projectName As String
    Dim invalidMessage As Integer

    'Check for valid page
    projectName = ActiveSheet.range(PROJECT_NAME_CELL).Value
    If Not IsValidJob(projectName) Then
        invalidMessage = MsgBox(NOT_PROJECT_PROMPT, vbOKOnly, NOT_PROJECT_TITLE)
        Exit Sub
    End If

    'Unprotect Sheets
    ActiveSheet.Unprotect
    Sheets(MASTER_SHEET_NAME).Unprotect

    'Unfilter
    With ActiveSheet
        If .AutoFilterMode Then
            If .FilterMode Then
                .ShowAllData
            End If
        Else
            If .FilterMode Then
                .ShowAllData
            End If
        End If
    End With
    With Sheets(MASTER_SHEET_NAME)
        If .AutoFilterMode Then
            If .FilterMode Then
                .ShowAllData
            End If
        Else
            If .FilterMode Then
                .ShowAllData
            End If
        End If
    End With

    'Sort Unit Parts List - Bldg, Part#, Hand
    ActiveSheet.range(UNIT_SORT_BEGIN & ":" & MAX_ROWS).Sort _
        Key1:=ActiveSheet.Columns(PROJECT_BUILDING_COLUMN), _
        Key2:=ActiveSheet.Columns(PROJECT_PART_NUM_COLUMN), _
        Key3:=ActiveSheet.Columns(PROJECT_HAND_COLUMN)

    'Sort Master By Job
    Sheets(MASTER_SHEET_NAME).range(MASTER_DATA_BEGIN & ":" & MAX_ROWS).Sort Key1:=Sheets(MASTER_SHEET_NAME).Columns(MASTER_PROJECT_COLUMN)

    'Delete Job From Master
    DeleteJobFromMaster (projectName)

    'Sort Master By Job
    Sheets(MASTER_SHEET_NAME).range(MASTER_DATA_BEGIN & ":" & MAX_ROWS).Sort Key1:=Sheets(MASTER_SHEET_NAME).Columns(MASTER_PROJECT_COLUMN)

    'Add Data to Master
    TransferDataFromUnitToMaster (projectName) 'Transfer Raw data.
    'Sort Least Significant
    Sheets(MASTER_SHEET_NAME).range(MASTER_DATA_BEGIN & ":" & MAX_ROWS).Sort _
        Key1:=Sheets(MASTER_SHEET_NAME).Columns(MASTER_PART_NUM_COLUMN), _
        Key2:=Sheets(MASTER_SHEET_NAME).Columns(MASTER_HAND_COLUMN), _
        Key3:=Sheets(MASTER_SHEET_NAME).Columns(MASTER_FLOOR_COLUMN)
    'Sort By building
    Sheets(MASTER_SHEET_NAME).range(MASTER_DATA_BEGIN & ":" & MAX_ROWS).Sort _
        Key1:=Sheets(MASTER_SHEET_NAME).Columns(MASTER_PROJECT_COLUMN), Order1:=xlDescending
    ConsolidateDuplicatesOnMaster (projectName) 'Find the entries that need combining
    Sheets(MASTER_SHEET_NAME).range(MASTER_DATA_BEGIN & ":" & MAX_ROWS).Sort _
        Key1:=Sheets(MASTER_SHEET_NAME).Columns(MASTER_PROJECT_COLUMN), Order1:=xlDescending, _
        Key2:=Sheets(MASTER_SHEET_NAME).Columns(MASTER_PART_NUM_COLUMN), Order1:=xlAscending 'Put back in order

    InsertDivisions

    'Costing
    InsertCosting

    'Sort Unit by Unit then part number
    ActiveSheet.range(PROJECT_DATA_BEGIN & ":" & MAX_ROWS).Sort _
        Key1:=ActiveSheet.Columns(PROJECT_PART_NUM_COLUMN), _
        Key2:=ActiveSheet.Columns(PROJECT_UNIT_COLUMN)

    'Protect Sheets
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
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

Function InsertDivisions()
    Dim masterRow As Integer: masterRow = MASTER_DATA_BEGIN
    Dim partNum As String: Dim j As Long
    While Not Sheets(MASTER_SHEET_NAME).Cells(masterRow, MASTER_PROJECT_COLUMN).Value = ""
        partNum = Sheets(MASTER_SHEET_NAME).Cells(masterRow, MASTER_PART_NUM_COLUMN).Value
        Sheets(MASTER_SHEET_NAME).Cells(masterRow, MASTER_DIVISION_COLUMN).Value = GetDivision(partNum) 'Set division
        masterRow = masterRow + 1
    Wend
End Function

Function ConsolidateDuplicatesOnMaster(projectName As String)
    Dim masterRow As Integer: masterRow = MASTER_DATA_BEGIN
    Dim qty As Long: Dim j As Integer
    While Not Sheets(MASTER_SHEET_NAME).Cells(masterRow, MASTER_PROJECT_COLUMN).Value = ""
        If Not Sheets(MASTER_SHEET_NAME).Cells(masterRow, MASTER_PROJECT_COLUMN).Value = projectName Then
            masterRow = masterRow + 1
        Else
            qty = 0
            j = GetEndOfSameBelowMaster(masterRow)
            For i = masterRow To j Step 1
                qty = qty + Sheets(MASTER_SHEET_NAME).Cells(i, MASTER_QUANTITY_COLUMN).Value 'Add up qty
            Next
            Sheets(MASTER_SHEET_NAME).Cells(masterRow, MASTER_QUANTITY_COLUMN).Value = qty 'Set first to correct quantity
            For i = masterRow + 1 To j Step 1
                Sheets(MASTER_SHEET_NAME).Cells(i, "A").EntireRow.Value = ""
            Next
            masterRow = j + 1
        End If
    Wend
End Function

Function TransferDataFromUnitToMaster(projectName As String)
    Dim bR As Long: Dim bL As Long: Dim f1R As Long:
    Dim f1L As Long: Dim f2R As Long: Dim f2L As Long:
    Dim f3R As Long: Dim f3L As Long: Dim f4R As Long:
    Dim f4L As Long: Dim gR As Long: Dim gL As Long
    Dim j As Integer: Dim hand As String:
    Dim building As String: Dim partNum As String: Dim measure As String
    Dim i As Integer: Dim multi As Integer
    Dim masterRow As Integer: Dim projectRow As Integer
    Dim right As String: Dim left As String
    masterRow = GetNextEmptyRow(MASTER_PROJECT_COLUMN, MASTER_DATA_BEGIN, MASTER_SHEET_NAME)
    projectRow = PROJECT_DATA_BEGIN
    While Not Cells(projectRow, PROJECT_PART_NUM_COLUMN).Value = ""
        bR = 0: bL = 0: f1R = 0: f1L = 0: f2R = 0: f2L = 0: f3R = 0: f3L = 0: f4R = 0: f4L = 0: gR = 0: gL = 0
        j = GetEndOfSameBelowUnit(projectRow)

        hand = GetHandUnit(j): building = GetBuildingUnit(j): partNum = GetPartNumUnit(j): measure = GetMeasureUnit(j)
        For i = projectRow To j Step 1
            multi = ActiveSheet.Cells(i, PROJECT_MULTIPLYER_COLUMN).Value
            bR = bR + GetBasementRight(hand, i) * multi
            bL = bL + GetBasementLeft(hand, i) * multi
            f1R = f1R + GetFirstRight(hand, i) * multi
            f1L = f1L + GetFirstLeft(hand, i) * multi
            f2R = f2R + GetSecondRight(hand, i) * multi
            f2L = f2L + GetSecondLeft(hand, i) * multi
            f3R = f3R + GetThirdRight(hand, i) * multi
            f3L = f3L + GetThirdLeft(hand, i) * multi
            f4R = f4R + GetFourthRight(hand, i) * multi
            f4L = f4L + GetFourthLeft(hand, i) * multi
            gR = gR + GetGeneralRight(hand, i) * multi
            gL = gL + GetGeneralLeft(hand, i) * multi
        Next
        'Determine Hand
        If hand = "" Then
            right = "": left = ""
        Else
            right = HAND_RIGHT: left = HAND_LEFT
        End If
        'Write Data
        'Basement
        masterRow = masterRow + WriteRowMaster(masterRow, projectName, partNum, right, bR, building, "B", measure)
        masterRow = masterRow + WriteRowMaster(masterRow, projectName, partNum, left, bL, building, "B", measure)
        'First
        masterRow = masterRow + WriteRowMaster(masterRow, projectName, partNum, right, f1R, building, "1", measure)
        masterRow = masterRow + WriteRowMaster(masterRow, projectName, partNum, left, f1L, building, "1", measure)
        'Second
        masterRow = masterRow + WriteRowMaster(masterRow, projectName, partNum, right, f2R, building, "2", measure)
        masterRow = masterRow + WriteRowMaster(masterRow, projectName, partNum, left, f2L, building, "2", measure)
        'Third
        masterRow = masterRow + WriteRowMaster(masterRow, projectName, partNum, right, f3R, building, "3", measure)
        masterRow = masterRow + WriteRowMaster(masterRow, projectName, partNum, left, f3L, building, "3", measure)
        'Fourth
        masterRow = masterRow + WriteRowMaster(masterRow, projectName, partNum, right, f4R, building, "4", measure)
        masterRow = masterRow + WriteRowMaster(masterRow, projectName, partNum, left, f4L, building, "4", measure)
        'General
        masterRow = masterRow + WriteRowMaster(masterRow, projectName, partNum, "", gR, building, "General", measure)
        masterRow = masterRow + WriteRowMaster(masterRow, projectName, partNum, "", gL, building, "General", measure)

        projectRow = j + 1
    Wend
End Function


Function WriteRowMaster(row As Integer, project As String, partNum As String, hand As String, qty As Long, bldg As String, floor As String, measure As String)
    If qty = 0 Then
        WriteRowMaster = 0
        Exit Function
    End If
    Sheets(MASTER_SHEET_NAME).Cells(row, MASTER_PROJECT_COLUMN).Value = project
    Sheets(MASTER_SHEET_NAME).Cells(row, MASTER_PART_NUM_COLUMN).Value = partNum
    Sheets(MASTER_SHEET_NAME).Cells(row, MASTER_HAND_COLUMN).Value = hand
    Sheets(MASTER_SHEET_NAME).Cells(row, MASTER_QUANTITY_COLUMN).Value = qty
    Sheets(MASTER_SHEET_NAME).Cells(row, MASTER_BUILDING_COLUMN).Value = bldg
    Sheets(MASTER_SHEET_NAME).Cells(row, MASTER_FLOOR_COLUMN).Value = floor
    Sheets(MASTER_SHEET_NAME).Cells(row, MASTER_MEASURE_COLUMN).Value = measure
    WriteRowMaster = 1
End Function

Function GetGeneralLeft(hand As String, row As Integer)
    If hand = HAND_RIGHT Then
        GetGeneralLeft = GetNum(row, UNIT_GENERAL_REV_COLUMN)
    ElseIf hand = HAND_LEFT Then
        GetGeneralLeft = GetNum(row, UNIT_GENERAL_STD_COLUMN)
    Else
        GetGeneralLeft = 0
    End If
End Function

Function GetGeneralRight(hand As String, row As Integer)
    If hand = HAND_RIGHT Then
        GetGeneralRight = GetNum(row, UNIT_GENERAL_STD_COLUMN)
    ElseIf hand = HAND_LEFT Then
        GetGeneralRight = GetNum(row, UNIT_GENERAL_REV_COLUMN)
    Else
        GetGeneralRight = GetNum(row, UNIT_GENERAL_STD_COLUMN) + GetNum(row, UNIT_GENERAL_REV_COLUMN)
    End If
End Function

Function GetFourthLeft(hand As String, row As Integer)
    If hand = HAND_RIGHT Then
        GetFourthLeft = GetNum(row, UNIT_FOURTH_REV_COLUMN)
    ElseIf hand = HAND_LEFT Then
        GetFourthLeft = GetNum(row, UNIT_FOURTH_STD_COLUMN)
    Else
        GetFourthLeft = 0
    End If
End Function

Function GetFourthRight(hand As String, row As Integer)
    If hand = HAND_RIGHT Then
        GetFourthRight = GetNum(row, UNIT_FOURTH_STD_COLUMN)
    ElseIf hand = HAND_LEFT Then
        GetFourthRight = GetNum(row, UNIT_FOURTH_REV_COLUMN)
    Else
        GetFourthRight = GetNum(row, UNIT_FOURTH_STD_COLUMN) + GetNum(row, UNIT_FOURTH_REV_COLUMN)
    End If
End Function

Function GetThirdLeft(hand As String, row As Integer)
    If hand = HAND_RIGHT Then
        GetThirdLeft = GetNum(row, UNIT_THIRD_REV_COLUMN)
    ElseIf hand = HAND_LEFT Then
        GetThirdLeft = GetNum(row, UNIT_THIRD_STD_COLUMN)
    Else
        GetThirdLeft = 0
    End If
End Function

Function GetThirdRight(hand As String, row As Integer)
    If hand = HAND_RIGHT Then
        GetThirdRight = GetNum(row, UNIT_THIRD_STD_COLUMN)
    ElseIf hand = HAND_LEFT Then
        GetThirdRight = GetNum(row, UNIT_THIRD_REV_COLUMN)
    Else
        GetThirdRight = GetNum(row, UNIT_THIRD_STD_COLUMN) + GetNum(row, UNIT_THIRD_REV_COLUMN)
    End If
End Function

Function GetSecondLeft(hand As String, row As Integer)
    If hand = HAND_RIGHT Then
        GetSecondLeft = GetNum(row, UNIT_SECOND_REV_COLUMN)
    ElseIf hand = HAND_LEFT Then
        GetSecondLeft = GetNum(row, UNIT_SECOND_STD_COLUMN)
    Else
        GetSecondLeft = 0
    End If
End Function

Function GetSecondRight(hand As String, row As Integer)
    If hand = HAND_RIGHT Then
        GetSecondRight = GetNum(row, UNIT_SECOND_STD_COLUMN)
    ElseIf hand = HAND_LEFT Then
        GetSecondRight = GetNum(row, UNIT_SECOND_REV_COLUMN)
    Else
        GetSecondRight = GetNum(row, UNIT_SECOND_STD_COLUMN) + GetNum(row, UNIT_SECOND_REV_COLUMN)
    End If
End Function

Function GetFirstLeft(hand As String, row As Integer)
    If hand = HAND_RIGHT Then
        GetFirstLeft = GetNum(row, UNIT_FIRST_REV_COLUMN)
    ElseIf hand = HAND_LEFT Then
        GetFirstLeft = GetNum(row, UNIT_FIRST_STD_COLUMN)
    Else
        GetFirstLeft = 0
    End If
End Function

Function GetFirstRight(hand As String, row As Integer)
    If hand = HAND_RIGHT Then
        GetFirstRight = GetNum(row, UNIT_FIRST_STD_COLUMN)
    ElseIf hand = HAND_LEFT Then
        GetFirstRight = GetNum(row, UNIT_FIRST_REV_COLUMN)
    Else
        GetFirstRight = GetNum(row, UNIT_FIRST_STD_COLUMN) + GetNum(row, UNIT_FIRST_REV_COLUMN)
    End If
End Function

Function GetBasementLeft(hand As String, row As Integer)
    If hand = HAND_RIGHT Then
        GetBasementLeft = GetNum(row, UNIT_BASEMENT_REV_COLUMN)
    ElseIf hand = HAND_LEFT Then
        GetBasementLeft = GetNum(row, UNIT_BASEMENT_STD_COLUMN)
    Else
        GetBasementLeft = 0
    End If
End Function

Function GetBasementRight(hand As String, row As Integer)
    If hand = HAND_RIGHT Then
        GetBasementRight = GetNum(row, UNIT_BASEMENT_STD_COLUMN)
    ElseIf hand = HAND_LEFT Then
        GetBasementRight = GetNum(row, UNIT_BASEMENT_REV_COLUMN)
    Else
        GetBasementRight = GetNum(row, UNIT_BASEMENT_STD_COLUMN) + GetNum(row, UNIT_BASEMENT_REV_COLUMN)
    End If
End Function

Function GetNum(row As Integer, col As String)
    If Cells(row, col).Value = "" Then
        GetNum = 0
    Else
        GetNum = Cells(row, col).Value
    End If
End Function
Function GetHandUnit(row As Integer)
    If ActiveSheet.Cells(row, PROJECT_HAND_COLUMN).Value = HAND_LEFT Or ActiveSheet.Cells(row, PROJECT_HAND_COLUMN).Value = HAND_RIGHT Then
        GetHandUnit = ActiveSheet.Cells(row, PROJECT_HAND_COLUMN).Value
    Else
        GetHandUnit = ""
    End If
End Function

Function GetPartNumUnit(row As Integer)
    GetPartNumUnit = ActiveSheet.Cells(row, PROJECT_PART_NUM_COLUMN).Value
End Function

Function GetBuildingUnit(row As Integer)
    GetBuildingUnit = ActiveSheet.Cells(row, PROJECT_BUILDING_COLUMN).Value
End Function

Function GetMeasureUnit(row As Integer)
    GetMeasureUnit = ActiveSheet.Cells(row, PROJECT_MEASURE_COLUMN).Value
End Function

Function GetNextEmptyRow(col As String, row As Integer, sheet As String)
    While Not Sheets(sheet).range(col & row).Value = ""
        row = row + 1
    Wend
    GetNextEmptyRow = row
End Function

Function DeleteJobFromMaster(projectName As String)
    For i = MAX_ROWS To MASTER_DATA_BEGIN Step -1
        If Sheets(MASTER_SHEET_NAME).Cells(i, MASTER_PROJECT_COLUMN).Value = projectName Then
            Sheets(MASTER_SHEET_NAME).Cells(i, MASTER_PROJECT_COLUMN).EntireRow.Delete
        End If
    Next
End Function

Function IsValidJob(projectName As String)
    If projectName = "" Then
        IsValidJob = False
        Exit Function
    End If
    For Each c In Sheets(VALID_SHEET_NAME).range(GetValidRangeDown(VALID_DATA_BEGIN, VALID_PROJECT_COLUMN, VALID_SHEET_NAME)).Cells
        If projectName = c.Value Then
            IsValidJob = True
            Exit Function
        End If
    Next
    IsValidJob = False
End Function

Function GetDivision(partNum As String)
    For Each c In Sheets(VALID_SHEET_NAME).range(GetValidRangeDown(VALID_DATA_BEGIN, VALID_DIVISION_COLUMN, VALID_SHEET_NAME)).Cells
        If left(partNum, 2) = left(c.Value, 2) Then
            GetDivision = c.Value
            Exit Function
        End If
    Next
    GetDivision = "No Division"
End Function

Function GetEndOfSameBelowUnit(rowNum As Integer)
    Dim i As Integer: Dim partN As String
    Dim hand As String: Dim bldg As String
    partN = ActiveSheet.Cells(rowNum, PROJECT_PART_NUM_COLUMN).Value
    hand = ActiveSheet.Cells(rowNum, PROJECT_HAND_COLUMN).Value
    bldg = ActiveSheet.Cells(rowNum, PROJECT_BUILDING_COLUMN).Value
    i = rowNum
    While ActiveSheet.Cells(i, PROJECT_PART_NUM_COLUMN).Value = partN And _
        ActiveSheet.Cells(i, PROJECT_HAND_COLUMN).Value = hand And ActiveSheet.Cells(i, PROJECT_BUILDING_COLUMN).Value = bldg
        i = i + 1
    Wend
    GetEndOfSameBelowUnit = (i - 1)
End Function

Function GetEndOfSameBelowMaster(rowNum As Integer)
    Dim i As Integer: Dim partN As String: Dim floor As String
    Dim hand As String: Dim bldg As String: Dim projectName As String: Dim measure As String
    partN = Sheets(MASTER_SHEET_NAME).Cells(rowNum, MASTER_PART_NUM_COLUMN).Value
    hand = Sheets(MASTER_SHEET_NAME).Cells(rowNum, MASTER_HAND_COLUMN).Value
    bldg = Sheets(MASTER_SHEET_NAME).Cells(rowNum, MASTER_BUILDING_COLUMN).Value
    projectName = Sheets(MASTER_SHEET_NAME).Cells(rowNum, MASTER_PROJECT_COLUMN).Value
    floor = Sheets(MASTER_SHEET_NAME).Cells(rowNum, MASTER_FLOOR_COLUMN).Value
    measure = Sheets(MASTER_SHEET_NAME).Cells(rowNum, MASTER_MEASURE_COLUMN).Value
    i = rowNum
    While Sheets(MASTER_SHEET_NAME).Cells(i, MASTER_PART_NUM_COLUMN).Value = partN And _
        Sheets(MASTER_SHEET_NAME).Cells(i, MASTER_HAND_COLUMN).Value = hand And Sheets(MASTER_SHEET_NAME).Cells(i, MASTER_BUILDING_COLUMN).Value = bldg And _
        Sheets(MASTER_SHEET_NAME).Cells(i, MASTER_PROJECT_COLUMN).Value = projectName And _
        Sheets(MASTER_SHEET_NAME).Cells(i, MASTER_FLOOR_COLUMN).Value = floor And Sheets(MASTER_SHEET_NAME).Cells(i, MASTER_MEASURE_COLUMN).Value = measure
        i = i + 1
    Wend
    GetEndOfSameBelowMaster = (i - 1)
End Function

Function GetValidRangeDown(row As Integer, col As String, sheet As String)
    Dim start As String
    start = col & row
    While Not Sheets(sheet).range(col & row).Value = ""
        row = row + 1
    Wend
    GetValidRangeDown = start & ":" & col & (row - 1)
End Function



