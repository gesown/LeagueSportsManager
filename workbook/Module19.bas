Attribute VB_Name = "Module19"
Sub Update_Player_Two()
Attribute Update_Player_Two.VB_ProcData.VB_Invoke_Func = "w\n14"
'
' Update_Player_Two Macro
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
    Range("B6:C7").Select
    Selection.Copy
    
    Sheets("Update").Select
    Range("D2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("Groups").Select
    Range("D6:D7").Select
    Application.CutCopyMode = False
    Selection.Copy
    
    Sheets("Update").Select
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("Groups").Select
    Range("Q6").Select
    Application.CutCopyMode = False
    Selection.Copy
    
    Sheets("Update").Select
    Range("E2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("Groups").Select
    Range("Q7").Select
    Application.CutCopyMode = False
    Selection.Copy
    
    Sheets("Update").Select
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("Groups").Select
    Range("R6:R7").Select
    Application.CutCopyMode = False
    Selection.Copy
    
    Sheets("Update").Select
    Range("I2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("Groups").Select
    Range("S6:S7").Select
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
    Range("A6:A7").Select
    ActiveCell.FormulaR1C1 = "DONE"
    Range("A8:A9").Select
    

End Sub

Sub Update_Player_Three()
Attribute Update_Player_Three.VB_ProcData.VB_Invoke_Func = "e\n14"
'
' Update_Player_Three Macro
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
    Range("B8:C9").Select
    
    Selection.Copy
    Sheets("Update").Select
    Range("D2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("D8:D9").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("Q8").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("E2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("Q9").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("R8:R9").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("I2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("S8:S9").Select
    
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
    
    Range("A8:A9").Select
    ActiveCell.FormulaR1C1 = "DONE"
    Range("A10:A11").Select
      

End Sub


Sub Update_Player_Four()
Attribute Update_Player_Four.VB_ProcData.VB_Invoke_Func = "r\n14"
'
' Update_Player_Four Macro
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
    Range("B10:C11").Select
    
    Selection.Copy
    Sheets("Update").Select
    Range("D2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("D10:D11").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("Q10").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("E2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("Q11").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("R10:R11").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("I2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("S10:S11").Select
    
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
    
    Range("A10:A11").Select
    ActiveCell.FormulaR1C1 = "DONE"
    Range("A12:A13").Select
    
   

End Sub
Sub Update_Player_Five()
Attribute Update_Player_Five.VB_ProcData.VB_Invoke_Func = "t\n14"
'
' Update_Player_Five Macro
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
    Range("B12:C13").Select
    
    Selection.Copy
    Sheets("Update").Select
    Range("D2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("D12:D13").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("Q12").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("E2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("Q13").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("R12:R13").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("I2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("S12:S13").Select
    
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
    
    Range("A12:A13").Select
    ActiveCell.FormulaR1C1 = "DONE"
    Range("A14:A15").Select
    
    

End Sub
Sub Update_Player_Six()
Attribute Update_Player_Six.VB_ProcData.VB_Invoke_Func = "y\n14"
'
' Update_Player_Six Macro
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
    Range("B14:C15").Select
    
    Selection.Copy
    Sheets("Update").Select
    Range("D2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("D14:D15").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("Q14").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("E2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("Q15").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("R14:R15").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("I2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("S14:S15").Select
    
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
    
    Range("A14:A15").Select
    ActiveCell.FormulaR1C1 = "DONE"
    Range("A16:A17").Select
    

    

End Sub
Sub Update_Player_Seven()
Attribute Update_Player_Seven.VB_ProcData.VB_Invoke_Func = "u\n14"
'
' Update_Player_Seven Macro
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
    Range("B16:C17").Select
    
    Selection.Copy
    Sheets("Update").Select
    Range("D2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("D16:D17").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("Q16").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("E2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("Q17").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("R16:R17").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("I2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("S16:S17").Select
    
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
    
    Range("A16:A17").Select
    ActiveCell.FormulaR1C1 = "DONE"
    Range("A18:A19").Select
    
    
    

End Sub
Sub Update_Player_Eight()
Attribute Update_Player_Eight.VB_ProcData.VB_Invoke_Func = "i\n14"
'
' Update_Player_Eight Macro
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
    Range("B18:C19").Select
    
    Selection.Copy
    Sheets("Update").Select
    Range("D2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("D18:D19").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("Q18").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("E2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("Q19").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("R18:R19").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("I2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("S18:S19").Select
    
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
    
    Range("A18:A19").Select
    ActiveCell.FormulaR1C1 = "DONE"
    Range("A20:A21").Select

    
    
    

End Sub

Sub Update_Player_Nine()
Attribute Update_Player_Nine.VB_ProcData.VB_Invoke_Func = "o\n14"
'
' Update_Player_Nine Macro
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
    Range("B20:C21").Select
    
    Selection.Copy
    Sheets("Update").Select
    Range("D2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("D20:D21").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("Q20").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("E2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("Q21").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("R20:R21").Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Update").Select
    Range("I2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Groups").Select
    Range("S20:S21").Select
    
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
    
    Range("A20:A21").Select
    ActiveCell.FormulaR1C1 = "DONE"
    Range("A22:A23").Select

    

End Sub
