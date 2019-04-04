Attribute VB_Name = "Module6"
Sub Paddle_1_Player()
'
' Paddle_1_Player Macro
' Adds paddle 1's player to the league list
'

'
    Application.ScreenUpdating = False
    Sheets("Search Function").Select
    Columns("E:H").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Home").Select
    Range("L26:O26").Select
    Selection.Copy
    Sheets("Search Function").Select
    Range("E2:H3001").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:A").Select
    Selection.Copy
    Columns("M:M").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(RC[3]:RC[1003])"
    Range("J2").Select
    Selection.Copy
    Range("J3:J3001").Select
    ActiveSheet.Paste
    Sheets("Home Player List Src").Select
    Columns("A:J").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Search Function").Select
    Columns("B:K").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:J").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("J2:J3001"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Largest to Smallest", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("B2:B3001"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Home Player List Src").Sort
        .SetRange Range("A1:J3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
    Range("L26").Select
    Selection.Copy
    Range("S21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            Range("S18").Select
    Selection.ClearContents
    Range("S18").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("G26:J26").Select
    Application.Run ("DeleteRefError")
    Application.ScreenUpdating = True
    
End Sub
Sub Paddle_2_Player()
'
' Paddle_2_Player Macro
' Adds paddle 2's player to the league list
'

'
    Application.ScreenUpdating = False
    Sheets("Search Function").Select
    Columns("E:H").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Home").Select
    Range("L27:O27").Select
    Selection.Copy
    Sheets("Search Function").Select
    Range("E2:H3001").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:A").Select
    Selection.Copy
    Columns("M:M").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(RC[3]:RC[1003])"
    Range("J2").Select
    Selection.Copy
    Range("J3:J3001").Select
    ActiveSheet.Paste
    Sheets("Home Player List Src").Select
    Columns("A:J").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Search Function").Select
    Columns("B:K").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:J").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("J2:J3001"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Largest to Smallest", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("B2:B3001"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Home Player List Src").Sort
        .SetRange Range("A1:J3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
    Range("L27").Select
    Selection.Copy
    Range("S21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            Range("S18").Select
    Selection.ClearContents
    Range("S18").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("G26:J26").Select
    Application.Run ("DeleteRefError")
    Application.ScreenUpdating = True
    
End Sub
Sub Paddle_3_Player()
Attribute Paddle_3_Player.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Paddle_3_Player Macro
' Adds paddle 3's player to the league list
'

'
    Application.ScreenUpdating = False
    Sheets("Search Function").Select
    Columns("E:H").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Home").Select
    Range("L28:O28").Select
    Selection.Copy
    Sheets("Search Function").Select
    Range("E2:H3001").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:A").Select
    Selection.Copy
    Columns("M:M").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(RC[3]:RC[1003])"
    Range("J2").Select
    Selection.Copy
    Range("J3:J3001").Select
    ActiveSheet.Paste
    Sheets("Home Player List Src").Select
    Columns("A:J").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Search Function").Select
    Columns("B:K").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:J").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("J2:J3001"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Largest to Smallest", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("B2:B3001"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Home Player List Src").Sort
        .SetRange Range("A1:J3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
    Range("L28").Select
    Selection.Copy
    Range("S21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            Range("S18").Select
    Selection.ClearContents
    Range("S18").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("G26:J26").Select
    Application.Run ("DeleteRefError")
    Application.ScreenUpdating = True
    
End Sub
Sub Paddle_4_Player()
Attribute Paddle_4_Player.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Paddle_4_Player Macro
' Adds paddle 4's player to the league list
'

'
    Application.ScreenUpdating = False
    Sheets("Search Function").Select
    Columns("E:H").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Home").Select
    Range("L29:O29").Select
    Selection.Copy
    Sheets("Search Function").Select
    Range("E2:H3001").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:A").Select
    Selection.Copy
    Columns("M:M").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(RC[3]:RC[1003])"
    Range("J2").Select
    Selection.Copy
    Range("J3:J3001").Select
    ActiveSheet.Paste
    Sheets("Home Player List Src").Select
    Columns("A:J").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Search Function").Select
    Columns("B:K").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:J").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("J2:J3001"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Largest to Smallest", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("B2:B3001"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Home Player List Src").Sort
        .SetRange Range("A1:J3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
    Range("L29").Select
    Selection.Copy
    Range("S21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            Range("S18").Select
    Selection.ClearContents
    Range("S18").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("G26:J26").Select
    Application.Run ("DeleteRefError")
    Application.ScreenUpdating = True
End Sub
Sub Paddle_5_Player()
Attribute Paddle_5_Player.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Paddle_5_Player Macro
' Adds paddle 5's player to the league list
'

'
    Application.ScreenUpdating = False
    Sheets("Search Function").Select
    Columns("E:H").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Home").Select
    Range("L30:O30").Select
    Selection.Copy
    Sheets("Search Function").Select
    Range("E2:H3001").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:A").Select
    Selection.Copy
    Columns("M:M").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(RC[3]:RC[1003])"
    Range("J2").Select
    Selection.Copy
    Range("J3:J3001").Select
    ActiveSheet.Paste
    Sheets("Home Player List Src").Select
    Columns("A:J").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Search Function").Select
    Columns("B:K").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:J").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("J2:J3001"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Largest to Smallest", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("B2:B3001"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Home Player List Src").Sort
        .SetRange Range("A1:J3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
    Range("L30").Select
    Selection.Copy
    Range("S21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            Range("S18").Select
    Selection.ClearContents
    Range("S18").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("G26:J26").Select
    Application.Run ("DeleteRefError")
    Application.ScreenUpdating = True
End Sub
Sub Paddle_6_Player()
Attribute Paddle_6_Player.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Paddle_3_Player Macro
' Adds paddle 3's player to the league list
'

'
    Application.ScreenUpdating = False
    Sheets("Search Function").Select
    Columns("E:H").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Home").Select
    Range("L31:O31").Select
    Selection.Copy
    Sheets("Search Function").Select
    Range("E2:H3001").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:A").Select
    Selection.Copy
    Columns("M:M").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(RC[3]:RC[1003])"
    Range("J2").Select
    Selection.Copy
    Range("J3:J3001").Select
    ActiveSheet.Paste
    Sheets("Home Player List Src").Select
    Columns("A:J").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Search Function").Select
    Columns("B:K").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:J").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("J2:J3001"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Largest to Smallest", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("B2:B3001"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Home Player List Src").Sort
        .SetRange Range("A1:J3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
        Range("L31").Select
    Selection.Copy
    Range("S21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            Range("S18").Select
    Selection.ClearContents
    Range("S18").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("G26:J26").Select
    Application.Run ("DeleteRefError")
    Application.ScreenUpdating = True
End Sub


Sub Paddle_7_Player()
Attribute Paddle_7_Player.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Paddle_7_Player Macro
' Adds paddle 7's player to the league list
'

'
    Application.ScreenUpdating = False
    Sheets("Search Function").Select
    Columns("E:H").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Home").Select
    Range("L32:O32").Select
    Selection.Copy
    Sheets("Search Function").Select
    Range("E2:H3001").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:A").Select
    Selection.Copy
    Columns("M:M").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(RC[3]:RC[1003])"
    Range("J2").Select
    Selection.Copy
    Range("J3:J3001").Select
    ActiveSheet.Paste
    Sheets("Home Player List Src").Select
    Columns("A:J").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Search Function").Select
    Columns("B:K").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:J").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("J2:J3001"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Largest to Smallest", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("B2:B3001"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Home Player List Src").Sort
        .SetRange Range("A1:J3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
    Range("L32").Select
    Selection.Copy
    Range("S21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            Range("S18").Select
    Selection.ClearContents
    Range("S18").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("G26:J26").Select
    Application.Run ("DeleteRefError")
    Application.ScreenUpdating = True
End Sub
Sub Paddle_8_Player()
Attribute Paddle_8_Player.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Paddle_8_Player Macro
' Adds paddle 8's player to the league list
'

'
    Application.ScreenUpdating = False
    Sheets("Search Function").Select
    Columns("E:H").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Home").Select
    Range("L33:O33").Select
    Selection.Copy
    Sheets("Search Function").Select
    Range("E2:H3001").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:A").Select
    Selection.Copy
    Columns("M:M").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(RC[3]:RC[1003])"
    Range("J2").Select
    Selection.Copy
    Range("J3:J3001").Select
    ActiveSheet.Paste
    Sheets("Home Player List Src").Select
    Columns("A:J").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Search Function").Select
    Columns("B:K").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:J").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("J2:J3001"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Largest to Smallest", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("B2:B3001"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Home Player List Src").Sort
        .SetRange Range("A1:J3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
    Range("L33").Select
    Selection.Copy
    Range("S21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            Range("S18").Select
    Selection.ClearContents
    Range("S18").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("G26:J26").Select
    Application.Run ("DeleteRefError")
    Application.ScreenUpdating = True
    
End Sub
Sub Paddle_9_Player()
Attribute Paddle_9_Player.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Paddle_9_Player Macro
' Adds paddle 9's player to the league list
'

'
    Application.ScreenUpdating = False
    Sheets("Search Function").Select
    Columns("E:H").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Home").Select
    Range("L34:O34").Select
    Selection.Copy
    Sheets("Search Function").Select
    Range("E2:H3001").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:A").Select
    Selection.Copy
    Columns("M:M").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(RC[3]:RC[1003])"
    Range("J2").Select
    Selection.Copy
    Range("J3:J3001").Select
    ActiveSheet.Paste
    Sheets("Home Player List Src").Select
    Columns("A:J").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Search Function").Select
    Columns("B:K").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:J").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("J2:J3001"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Largest to Smallest", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("B2:B3001"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Home Player List Src").Sort
        .SetRange Range("A1:J3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
    Range("L34").Select
    Selection.Copy
    Range("S21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            Range("S18").Select
    Selection.ClearContents
    Range("S18").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("G26:J26").Select
    Application.Run ("DeleteRefError")
    Application.ScreenUpdating = True
    
End Sub
Sub Paddle_10_Player()
Attribute Paddle_10_Player.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Paddle_10_Player Macro
' Adds paddle 10's player to the league list
'

'
    Application.ScreenUpdating = False
    Sheets("Search Function").Select
    Columns("E:H").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Home").Select
    Range("L35:O35").Select
    Selection.Copy
    Sheets("Search Function").Select
    Range("E2:H3001").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:A").Select
    Selection.Copy
    Columns("M:M").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(RC[3]:RC[1003])"
    Range("J2").Select
    Selection.Copy
    Range("J3:J3001").Select
    ActiveSheet.Paste
    Sheets("Home Player List Src").Select
    Columns("A:J").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Search Function").Select
    Columns("B:K").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:J").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("J2:J3001"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Largest to Smallest", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("B2:B3001"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Home Player List Src").Sort
        .SetRange Range("A1:J3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
    Range("L35").Select
    Selection.Copy
    Range("S21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            Range("S18").Select
    Selection.ClearContents
    Range("S18").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("G26:J26").Select
    Application.Run ("DeleteRefError")
    Application.ScreenUpdating = True
    
    
End Sub
Sub Paddle_11_Player()
Attribute Paddle_11_Player.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Paddle_11_Player Macro
' Adds paddle 11's player to the league list
'

'
    Application.ScreenUpdating = False
    Sheets("Search Function").Select
    Columns("E:H").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Home").Select
    Range("L36:O36").Select
    Selection.Copy
    Sheets("Search Function").Select
    Range("E2:H3001").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:A").Select
    Selection.Copy
    Columns("M:M").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(RC[3]:RC[1003])"
    Range("J2").Select
    Selection.Copy
    Range("J3:J3001").Select
    ActiveSheet.Paste
    Sheets("Home Player List Src").Select
    Columns("A:J").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Search Function").Select
    Columns("B:K").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:J").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("J2:J3001"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Largest to Smallest", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("B2:B3001"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Home Player List Src").Sort
        .SetRange Range("A1:J3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
    Range("L36").Select
    Selection.Copy
    Range("S21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("S18").Select
    Selection.ClearContents
    Range("S18").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("G26:J26").Select
    Application.Run ("DeleteRefError")
    Application.ScreenUpdating = True
    
End Sub
Sub Delete_1st_row_Player()
Attribute Delete_1st_row_Player.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Delete_1st_row_Player Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Search Function").Select
    Columns("E:H").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Home").Select
    Range("L26:O26").Select
    Selection.Copy
    Sheets("Search Function").Select
    Range("E2:H3001").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("G2").Select
    ActiveCell.FormulaR1C1 = "3500"
    Range("G2").Select
    Selection.Copy
    Range("G2:G3001").Select
    ActiveSheet.Paste
    Columns("A:A").Select
    Selection.Copy
    Columns("M:M").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(RC[3]:RC[1003])"
    Range("J2").Select
    Selection.Copy
    Range("J3:J3001").Select
    ActiveSheet.Paste
    Sheets("Home Player List Src").Select
    Columns("A:J").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Search Function").Select
    Columns("B:K").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:J").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("J2:J3001"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Largest to Smallest", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("B2:B3001"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Home Player List Src").Sort
        .SetRange Range("A1:J3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
    Range("L26").Select
    Selection.Copy
    Range("S21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            Range("S18").Select
    Selection.ClearContents
    Range("S18").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("G26:J26").Select
    ActiveWindow.SmallScroll Down:=1
    Application.ScreenUpdating = True

End Sub
Sub Delete_2nd_row_Player()
Attribute Delete_2nd_row_Player.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Delete_2nd_row_Player Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Search Function").Select
    Columns("E:H").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Home").Select
    Range("L27:O27").Select
    Selection.Copy
    Sheets("Search Function").Select
    Range("E2:H3001").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("G2").Select
    ActiveCell.FormulaR1C1 = "3500"
    Range("G2").Select
    Selection.Copy
    Range("G2:G3001").Select
    ActiveSheet.Paste
    Columns("A:A").Select
    Selection.Copy
    Columns("M:M").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(RC[3]:RC[1003])"
    Range("J2").Select
    Selection.Copy
    Range("J3:J3001").Select
    ActiveSheet.Paste
    Sheets("Home Player List Src").Select
    Columns("A:J").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Search Function").Select
    Columns("B:K").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:J").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("J2:J3001"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Largest to Smallest", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("B2:B3001"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Home Player List Src").Sort
        .SetRange Range("A1:J3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
    Range("L27").Select
    Selection.Copy
    Range("S21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            Range("S18").Select
    Selection.ClearContents
    Range("S18").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("G26:J26").Select
    ActiveWindow.SmallScroll Down:=1
    Application.ScreenUpdating = True
End Sub
Sub Delete_3rd_row_Player()
Attribute Delete_3rd_row_Player.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Delete_3rd_row_Player Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Search Function").Select
    Columns("E:H").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Home").Select
    Range("L28:O28").Select
    Selection.Copy
    Sheets("Search Function").Select
    Range("E2:H3001").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("G2").Select
    ActiveCell.FormulaR1C1 = "3500"
    Range("G2").Select
    Selection.Copy
    Range("G2:G3001").Select
    ActiveSheet.Paste
    Columns("A:A").Select
    Selection.Copy
    Columns("M:M").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(RC[3]:RC[1003])"
    Range("J2").Select
    Selection.Copy
    Range("J3:J3001").Select
    ActiveSheet.Paste
    Sheets("Home Player List Src").Select
    Columns("A:J").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Search Function").Select
    Columns("B:K").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:J").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("J2:J3001"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Largest to Smallest", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("B2:B3001"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Home Player List Src").Sort
        .SetRange Range("A1:J3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
    Range("L28").Select
    Selection.Copy
    Range("S21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            Range("S18").Select
    Selection.ClearContents
    Range("S18").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("G26:J26").Select
    ActiveWindow.SmallScroll Down:=1
    Application.ScreenUpdating = True

End Sub
Sub Delete_4th_row_Player()
Attribute Delete_4th_row_Player.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Delete_4th_row_Player Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Search Function").Select
    Columns("E:H").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Home").Select
    Range("L29:O29").Select
    Selection.Copy
    Sheets("Search Function").Select
    Range("E2:H3001").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("G2").Select
    ActiveCell.FormulaR1C1 = "3500"
    Range("G2").Select
    Selection.Copy
    Range("G2:G3001").Select
    ActiveSheet.Paste
    Columns("A:A").Select
    Selection.Copy
    Columns("M:M").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(RC[3]:RC[1003])"
    Range("J2").Select
    Selection.Copy
    Range("J3:J3001").Select
    ActiveSheet.Paste
    Sheets("Home Player List Src").Select
    Columns("A:J").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Search Function").Select
    Columns("B:K").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:J").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("J2:J3001"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Largest to Smallest", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("B2:B3001"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Home Player List Src").Sort
        .SetRange Range("A1:J3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
    Range("L29").Select
    Selection.Copy
    Range("S21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            Range("S18").Select
    Selection.ClearContents
    Range("S18").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("G26:J26").Select
    ActiveWindow.SmallScroll Down:=1
    Application.ScreenUpdating = True
End Sub
Sub Delete_5th_row_Player()
Attribute Delete_5th_row_Player.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Delete_5th_row_Player Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Search Function").Select
    Columns("E:H").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Home").Select
    Range("L30:O30").Select
    Selection.Copy
    Sheets("Search Function").Select
    Range("E2:H3001").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("G2").Select
    ActiveCell.FormulaR1C1 = "3500"
    Range("G2").Select
    Selection.Copy
    Range("G2:G3001").Select
    ActiveSheet.Paste
    Columns("A:A").Select
    Selection.Copy
    Columns("M:M").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(RC[3]:RC[1003])"
    Range("J2").Select
    Selection.Copy
    Range("J3:J3001").Select
    ActiveSheet.Paste
    Sheets("Home Player List Src").Select
    Columns("A:J").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Search Function").Select
    Columns("B:K").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:J").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("J2:J3001"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Largest to Smallest", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("B2:B3001"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Home Player List Src").Sort
        .SetRange Range("A1:J3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
    Range("L30").Select
    Selection.Copy
    Range("S21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            Range("S18").Select
    Selection.ClearContents
    Range("S18").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("G26:J26").Select
    ActiveWindow.SmallScroll Down:=1
    Application.ScreenUpdating = True
End Sub
Sub Delete_6th_row_Player()
Attribute Delete_6th_row_Player.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Delete_6th_row_Player Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Search Function").Select
    Columns("E:H").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Home").Select
    Range("L31:O31").Select
    Selection.Copy
    Sheets("Search Function").Select
    Range("E2:H3001").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("G2").Select
    ActiveCell.FormulaR1C1 = "3500"
    Range("G2").Select
    Selection.Copy
    Range("G2:G3001").Select
    ActiveSheet.Paste
    Columns("A:A").Select
    Selection.Copy
    Columns("M:M").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(RC[3]:RC[1003])"
    Range("J2").Select
    Selection.Copy
    Range("J3:J3001").Select
    ActiveSheet.Paste
    Sheets("Home Player List Src").Select
    Columns("A:J").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Search Function").Select
    Columns("B:K").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:J").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("J2:J3001"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Largest to Smallest", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("B2:B3001"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Home Player List Src").Sort
        .SetRange Range("A1:J3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
    Range("L31").Select
    Selection.Copy
    Range("S21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            Range("S18").Select
    Selection.ClearContents
    Range("S18").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("G26:J26").Select
    ActiveWindow.SmallScroll Down:=1
    Application.ScreenUpdating = True
End Sub
Sub Delete_7th_row_Player()
Attribute Delete_7th_row_Player.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Delete_7th_row_Player Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Search Function").Select
    Columns("E:H").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Home").Select
    Range("L32:O32").Select
    Selection.Copy
    Sheets("Search Function").Select
    Range("E2:H3001").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("G2").Select
    ActiveCell.FormulaR1C1 = "3500"
    Range("G2").Select
    Selection.Copy
    Range("G2:G3001").Select
    ActiveSheet.Paste
    Columns("A:A").Select
    Selection.Copy
    Columns("M:M").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(RC[3]:RC[1003])"
    Range("J2").Select
    Selection.Copy
    Range("J3:J3001").Select
    ActiveSheet.Paste
    Sheets("Home Player List Src").Select
    Columns("A:J").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Search Function").Select
    Columns("B:K").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:J").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("J2:J3001"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Largest to Smallest", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("B2:B3001"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Home Player List Src").Sort
        .SetRange Range("A1:J3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
    Range("L32").Select
    Selection.Copy
    Range("S21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            Range("S18").Select
    Selection.ClearContents
    Range("S18").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("G26:J26").Select
    ActiveWindow.SmallScroll Down:=1
    Application.ScreenUpdating = True
End Sub
Sub Delete_8th_row_Player()
Attribute Delete_8th_row_Player.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Delete_8th_row_Player Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Search Function").Select
    Columns("E:H").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Home").Select
    Range("L33:O33").Select
    Selection.Copy
    Sheets("Search Function").Select
    Range("E2:H3001").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("G2").Select
    ActiveCell.FormulaR1C1 = "3500"
    Range("G2").Select
    Selection.Copy
    Range("G2:G3001").Select
    ActiveSheet.Paste
    Columns("A:A").Select
    Selection.Copy
    Columns("M:M").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(RC[3]:RC[1003])"
    Range("J2").Select
    Selection.Copy
    Range("J3:J3001").Select
    ActiveSheet.Paste
    Sheets("Home Player List Src").Select
    Columns("A:J").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Search Function").Select
    Columns("B:K").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:J").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("J2:J3001"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Largest to Smallest", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("B2:B3001"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Home Player List Src").Sort
        .SetRange Range("A1:J3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
    Range("L33").Select
    Selection.Copy
    Range("S21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            Range("S18").Select
    Selection.ClearContents
    Range("S18").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("G26:J26").Select
    ActiveWindow.SmallScroll Down:=1
    Application.ScreenUpdating = True
End Sub
Sub Delete_9th_row_Player()
Attribute Delete_9th_row_Player.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Delete_9th_row_Player Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Search Function").Select
    Columns("E:H").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Home").Select
    Range("L34:O34").Select
    Selection.Copy
    Sheets("Search Function").Select
    Range("E2:H3001").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("G2").Select
    ActiveCell.FormulaR1C1 = "3500"
    Range("G2").Select
    Selection.Copy
    Range("G2:G3001").Select
    ActiveSheet.Paste
    Columns("A:A").Select
    Selection.Copy
    Columns("M:M").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(RC[3]:RC[1003])"
    Range("J2").Select
    Selection.Copy
    Range("J3:J3001").Select
    ActiveSheet.Paste
    Sheets("Home Player List Src").Select
    Columns("A:J").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Search Function").Select
    Columns("B:K").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:J").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("J2:J3001"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Largest to Smallest", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("B2:B3001"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Home Player List Src").Sort
        .SetRange Range("A1:J3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
    Range("L34").Select
    Selection.Copy
    Range("S21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            Range("S18").Select
    Selection.ClearContents
    Range("S18").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("G26:J26").Select
    ActiveWindow.SmallScroll Down:=1
    Application.ScreenUpdating = True
End Sub
Sub Delete_10th_row_Player()
Attribute Delete_10th_row_Player.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Delete_10th_row_Player Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Search Function").Select
    Columns("E:H").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Home").Select
    Range("L35:O35").Select
    Selection.Copy
    Sheets("Search Function").Select
    Range("E2:H3001").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("G2").Select
    ActiveCell.FormulaR1C1 = "3500"
    Range("G2").Select
    Selection.Copy
    Range("G2:G3001").Select
    ActiveSheet.Paste
    Columns("A:A").Select
    Selection.Copy
    Columns("M:M").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(RC[3]:RC[1003])"
    Range("J2").Select
    Selection.Copy
    Range("J3:J3001").Select
    ActiveSheet.Paste
    Sheets("Home Player List Src").Select
    Columns("A:J").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Search Function").Select
    Columns("B:K").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:J").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("J2:J3001"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Largest to Smallest", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("B2:B3001"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Home Player List Src").Sort
        .SetRange Range("A1:J3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
    Range("L35").Select
    Selection.Copy
    Range("S21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            Range("S18").Select
    Selection.ClearContents
    Range("S18").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("G26:J26").Select
    ActiveWindow.SmallScroll Down:=1
    Application.ScreenUpdating = True
    
End Sub
Sub Delete_11the_row_Player()
Attribute Delete_11the_row_Player.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Delete_11the_row_Player Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Search Function").Select
    Columns("E:H").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Home").Select
    Range("L36:O36").Select
    Selection.Copy
    Sheets("Search Function").Select
    Range("E2:H3001").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("G2").Select
    ActiveCell.FormulaR1C1 = "3500"
    Range("G2").Select
    Selection.Copy
    Range("G2:G3001").Select
    ActiveSheet.Paste
    Columns("A:A").Select
    Selection.Copy
    Columns("M:M").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(RC[3]:RC[1003])"
    Range("J2").Select
    Selection.Copy
    Range("J3:J3001").Select
    ActiveSheet.Paste
    Sheets("Home Player List Src").Select
    Columns("A:J").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Search Function").Select
    Columns("B:K").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:J").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("J2:J3001"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
        :="Largest to Smallest", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("B2:B3001"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Home Player List Src").Sort
        .SetRange Range("A1:J3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
    Range("L36").Select
    Selection.Copy
    Range("S21").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
            Range("S18").Select
    Selection.ClearContents
    Range("S18").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("G26:J26").Select
    ActiveWindow.SmallScroll Down:=1
    Application.ScreenUpdating = True
End Sub
