Attribute VB_Name = "Module29"

Sub Go_To_Home()
Attribute Go_To_Home.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Go_To_Home Macro
'

'
    Sheets("Home").Select
    Range("A1").Select
End Sub
Sub Go_To_Update_Player()
Attribute Go_To_Update_Player.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Go_To_Update_Player Macro
'

'
    Sheets("Adjust-Delete").Select
    Range("E13").Select
End Sub
Sub Start_Update_Delete_Player()
Attribute Start_Update_Delete_Player.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Start_Update_Delete_Player Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Upd-Del-Plyr-List").Select
    Rows("2:19").Select
    Application.GoTo Reference:="R2:R3000"
    Selection.ClearContents
    Sheets("Player Archive").Select
    Application.GoTo Reference:="R2:R3000"
    Selection.Copy
    Sheets("Upd-Del-Plyr-List").Select
    Application.GoTo Reference:="R2:R3000"
    ActiveSheet.Paste
    Columns("A:U").Select
    ActiveWorkbook.Worksheets("Upd-Del-Plyr-List").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Upd-Del-Plyr-List").Sort.SortFields.Add Key:=Range _
        ("E2:E3000"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Upd-Del-Plyr-List").Sort
        .SetRange Range("A1:U3000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Adjust-Delete").Select
    Range("E13").Select
    Application.ScreenUpdating = True
    
End Sub

Sub Look_Up_Plyr_By_Rating()
Attribute Look_Up_Plyr_By_Rating.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Look_Up_Plyr_By_Rating Macro
'

'
    Application.ScreenUpdating = False
    Range("U2").Select
    Selection.Copy
    Range("U1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets("Upd-Del-Plyr-List").Select
    Columns("A:U").Select
    ActiveWorkbook.Worksheets("Upd-Del-Plyr-List").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Upd-Del-Plyr-List").Sort.SortFields.Add Key:=Range _
        ("E2:E3000"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Upd-Del-Plyr-List").Sort
        .SetRange Range("A1:U3000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Adjust-Delete").Select
    Range("E13").Select
    Application.ScreenUpdating = True
End Sub
Sub Look_Up_Plyr_By_Name()
Attribute Look_Up_Plyr_By_Name.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Look_Up_Plyr_By_Name Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Upd-Del-Plyr-List").Select
    Columns("A:U").Select
    ActiveWorkbook.Worksheets("Upd-Del-Plyr-List").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Upd-Del-Plyr-List").Sort.SortFields.Add Key:=Range _
        ("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Upd-Del-Plyr-List").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Upd-Del-Plyr-List").Sort.SortFields.Add Key:=Range _
        ("D2:D3000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Upd-Del-Plyr-List").Sort.SortFields.Add Key:=Range _
        ("E2:E3000"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Upd-Del-Plyr-List").Sort
        .SetRange Range("A1:U3000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("V2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(COUNT(SEARCH('Adjust-Delete'!R13C5,RC[-18])),""Yes"","""")"
    Range("V2").Select
    Selection.Copy
    Application.GoTo Reference:="R2C22:R3000C22"
    ActiveSheet.Paste
    Range("V1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Range("V2").Select
    Sheets("Adjust-Delete").Select
    Range("U5").Select
    Selection.Copy
    Range("U1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("E13").Select
    Application.ScreenUpdating = True
End Sub
