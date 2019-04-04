Attribute VB_Name = "Module18"
Sub Update_Player_One()
Attribute Update_Player_One.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' Update_Player_One Macro
'
'
'

    Sheets("Player Archive").Select
    Columns("T:U").Select
    Selection.ClearContents
    Sheets("OneTo3000").Select
    Columns("A:A").Select
    Selection.Copy
    Sheets("Player Archive").Select
    Columns("U:U").Select
    ActiveSheet.Paste
    Columns("T:T").Select
    ActiveSheet.Paste
    Sheets("Update").Select
    Rows("2").Select
    Selection.ClearContents
    
    Sheets("Groups").Select
    Range("B4:C5").Select
    
    Selection.Copy
    Sheets("Update").Select
    Range("D2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("D4:D5").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("Q4").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("E2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("Q5").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("R4:R5").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("I2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("S4:S5").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("J2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveWindow.SmallScroll ToRight:=-1
    Range("J6").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Range("J7").Select
    Sheets("Player Archive").Select
    Rows("60:60").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("Update").Select
    Rows("5:5").Select
    Selection.Copy
    Sheets("Player Archive").Select
    Rows("60:60").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A59:B59").Select
    Range("B59").Activate
    Application.CutCopyMode = False
    Selection.Copy
    Range("A60:B60").Select
    ActiveSheet.Paste
    Range("M59").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M60").Select
    ActiveSheet.Paste
    Range("Q60").Select
    Application.CutCopyMode = False
    Selection.NumberFormat = "mm/dd/yy;@"
    Columns("A:U").Select
    ActiveWorkbook.Worksheets("Player Archive").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Player Archive").Sort.SortFields.Add Key:=Range( _
        "D2:D3002"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Player Archive").Sort.SortFields.Add Key:=Range( _
        "U2:U3002"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Player Archive").Sort
        .SetRange Range("A1:U3002")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveSheet.Range("$A$1:$U$3002").RemoveDuplicates Columns:=4, Header:=xlNo
    ActiveWorkbook.Worksheets("Player Archive").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Player Archive").Sort.SortFields.Add Key:=Range( _
        "U2:U3002"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Player Archive").Sort
        .SetRange Range("A1:U3002")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Attendance").Select
    Range("A60").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A61").Select
    Selection.Copy
    Range("A60").Select
    ActiveSheet.Paste
    Range("A61").Select
    Sheets("Search Function").Select
    Range("A60:I60").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A59:I59").Select
    Selection.Copy
    Range("A60:I60").Select
    ActiveSheet.Paste
    Range("A59:I59").Select
    
    Sheets("Groups").Select
    Range("A4:A5").Select
    ActiveCell.FormulaR1C1 = "DONE"
    Range("A6:A7").Select
      

End Sub

