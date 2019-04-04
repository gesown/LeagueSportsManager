Attribute VB_Name = "Module17"
Sub Add_New_Player()
Attribute Add_New_Player.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Add_New_Player Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Printable Results").Select
    Application.Run ("FilterOFF_ForPrintableResults")
    Sheets("Rankings").Select
    Application.Run ("FilterOFF_ForRankings")
    Sheets("Player Archive").Select
    Rows("60:60").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A59:B59").Select
    Selection.Copy
    Range("A60:B60").Select
    ActiveSheet.Paste
    Range("M59").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M60").Select
    ActiveSheet.Paste
    Range("T2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("T3").Formula = "=T2+1"
    Range("T4").Select
    Range("T3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("T4:T3001").Select
    ActiveSheet.Paste
    Sheets("Home").Select
    Range("F16:H16").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Player Archive").Select
    Range("D60").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Home").Select
    Range("J16:K16").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Player Archive").Select
    Range("E60").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("E60").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("F60").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("G60").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "0"
    Range("H60").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("I60").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("J60").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("K60").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("L60").Select
    ActiveCell.FormulaR1C1 = "0.1"
    Range("N60").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("O60").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("P60").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("A60").Select
    Sheets("Attendance").Select
    Rows("60:60").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A59").Select
    Selection.Copy
    Range("A60").Select
    ActiveSheet.Paste
    Range("B59").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B60").Select
    ActiveSheet.Paste
    Sheets("Search Function").Select
    Rows("60:60").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A59:C59").Select
    Selection.Copy
    Range("A60:C60").Select
    ActiveSheet.Paste
    Range("J59:K59").Select
    Selection.Copy
    Range("J60:K60").Select
    ActiveSheet.Paste
    Sheets("Player Archive").Select
    Cells.Select
    Range("A50").Activate
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Players").Select
    Cells.Select
    ActiveSheet.Paste
    Columns("A:R").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Players").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Players").Sort.SortFields.Add Key:=Range( _
        "E2:E3004"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Players").Sort
        .SetRange Range("A1:R3004")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Sheets("Season Groups").Select
    Application.GoTo Reference:="R60"
    Rows("60:60").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B59:E59").Select
    Selection.Copy
    Range("B60").Select
    Range("B59:E3176").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Sheets("Home").Select
    Range("F16:H16").Select
    Selection.ClearContents
    Range("J16:K16").Select
    Selection.ClearContents
    Range("M16:O16").Select
    Selection.ClearContents
    Range("F16:H16").Select
    ActiveCell.FormulaR1C1 = "Player Added"
    Range("F16:H16").Select
    Application.ScreenUpdating = True
End Sub
