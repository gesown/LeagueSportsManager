Attribute VB_Name = "Module24"
Sub Master_Season_Reset()
'
' Master_Season_Reset Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Player Archive").Select
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("G2").Select
    Selection.Copy
    Application.GoTo Reference:="R2C7:R3178C12"
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveWindow.SmallScroll Down:=-15
    Application.GoTo Reference:="R2C14:R3178C14"
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveWindow.SmallScroll Down:=-15
    Application.GoTo Reference:="R2C16:R3178C16"
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveWindow.SmallScroll Down:=-15
    Columns("D:D").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Selection.EntireRow.Delete
    Range("A1").Select
    Sheets("Printable Results").Select
    Range("A1:A2").Select
    Application.GoTo Reference:="R2C7:R3178C12"
    Application.GoTo Reference:="R7C7:R3178C12"
    ActiveWindow.SmallScroll Down:=-12
    Range("G7").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("G7").Select
    Selection.Copy
    Application.GoTo Reference:="R7C7:R3178C12"
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.GoTo Reference:="R7C14:R3178C14"
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.GoTo Reference:="R7C16:R3178C16"
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveWindow.SmallScroll Down:=-21
    Application.GoTo Reference:="R7C4:R3300C4"
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Selection.EntireRow.Delete
    Range("A1:A2").Select
    ActiveWindow.ScrollWorkbookTabs Position:=xlLast
    Sheets("Attendance").Select
    Columns("D:ZZ").Select
    Selection.Delete Shift:=xlToLeft
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[1]:RC[700])"
    Range("B2").Select
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets("Up Down Arrows").Select
    ActiveWindow.ScrollWorkbookTabs Position:=xlFirst
    Sheets("Players").Select
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 1
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("G2").Select
    Selection.Copy
    Application.GoTo Reference:="R2C7:R3178C12"
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.GoTo Reference:="R2C14:R3178C14"
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveWindow.SmallScroll Down:=-18
    Application.GoTo Reference:="R2C16:R3178C16"
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("G2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "0"
    Range("A1").Select
    Sheets("Score Cards").Select
    Range("A1").Select
    Sheets("Home").Select
    Range("S21").Select
    Selection.ClearContents
    Range("G26:J26").Select
    Selection.ClearContents
    Range("F16:H16").Select
    Selection.ClearContents
    Range("D16").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A16:C16").Select
    Range("A16:C16").Select
        ActiveWindow.ScrollWorkbookTabs Position:=xlLast
    Sheets("Season Groups").Select
    Columns("G:BX").Select
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 1
    Selection.Delete Shift:=xlToLeft
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(RC[2]:RC[22])"
    Range("E2").Select
    Selection.Copy
    Range("E2:E3164").Select
    ActiveSheet.Paste
    Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Range("A1").Select
    Sheets("Home").Select
    Range("A16:C16").Select
    ActiveCell.FormulaR1C1 = "MASTER RESET COMPLETE"
    Application.ScreenUpdating = True
End Sub

Sub Calculate_Season_Data()
'
' Calculate_Season_Data Macro
'

'

    Application.ScreenUpdating = False
    ActiveWindow.ScrollWorkbookTabs Position:=xlLast
    Sheets("End Of Season Data").Select
    Rows("1:1001").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.ScrollWorkbookTabs Position:=xlFirst

    Sheets("Printable Results").Select
    Range("D6:P306").Select
    Selection.Copy
    ActiveWindow.ScrollWorkbookTabs Position:=xlLast
    Sheets("End Of Season Data").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:M").Select
    Columns("A:M").EntireColumn.AutoFit
    Application.CutCopyMode = False
    Columns("C:C").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("J:J").Select
    Selection.NumberFormat = "0.00%"
    Columns("A:L").Select
    ActiveWorkbook.Worksheets("End Of Season Data").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("End Of Season Data").Sort.SortFields.Add Key:= _
        Range("J2:J314"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("End Of Season Data").Sort
        .SetRange Range("A1:L314")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<13,""*_**_*"",RC[-1])"
    Range("M2").Select
    Selection.Copy
    Range("M3:M305").Select
    ActiveSheet.Paste
    Columns("M:M").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("M:M").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("M1").Select
    Application.CutCopyMode = False
    Columns("M:M").Select
    Selection.Replace What:="*_**_*", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("M2:M242").Select
    ActiveWindow.SmallScroll Down:=66
    Range("M2:M305").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    Rows("6:305").Select
    Selection.ClearContents
    Range("A1:L1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("M2:M5").Select
    Selection.ClearContents
    Range("J2:J5").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    Range("A7").Select
    ActiveWindow.ScrollWorkbookTabs Position:=xlFirst
    Sheets("Printable Results").Select
    Range("D6:P306").Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveWindow.ScrollWorkbookTabs Position:=xlLast
    Sheets("End Of Season Data").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A7:L305").Select
    ActiveWorkbook.Worksheets("End Of Season Data").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("End Of Season Data").Sort.SortFields.Add Key:= _
        Range("K8:K305"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("End Of Season Data").Sort
        .SetRange Range("A7:L305")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("M8").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<13,""*_**_*"",RC[-1])"
    Range("M8").Select
    Selection.Copy
    Range("M9:M305").Select
    ActiveSheet.Paste
    Range("M8:M305").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M8:M305").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("M7").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Range("M8:M305").Select
    Selection.Replace What:="*_**_*", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    Range("A13:M305").Select
    Selection.ClearContents
    Range("A7:L7").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("M8:M12").Select
    Selection.ClearContents
    Range("K8:K12").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    ActiveWindow.ScrollWorkbookTabs Position:=xlFirst
    Sheets("Printable Results").Select
    Range("D6:P306").Select
    Selection.Copy
    ActiveWindow.ScrollWorkbookTabs Position:=xlLast
    Sheets("End Of Season Data").Select
    Range("A14").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("L:L").Select
    Range("L9").Activate
    Selection.Replace What:="#REF!", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("A14:L305").Select
    ActiveWorkbook.Worksheets("End Of Season Data").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("End Of Season Data").Sort.SortFields.Add Key:= _
        Range("L15:L305"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("End Of Season Data").Sort
        .SetRange Range("A14:L305")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("M15").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<8,""*_**_*"",RC[-1])"
    Range("M15").Select
    Selection.Copy
    Range("M16:M305").Select
    ActiveSheet.Paste
    Range("M15:M305").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M15:M305").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("M15:M305").Select
    Selection.Replace What:="*_**_*", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("M15:M305").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Selection.EntireRow.Delete
    
    Range("A14:L14").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10092441
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("L15:L305").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Columns("L:L").EntireColumn.AutoFit
    Range("M1:M305").Select
    Selection.ClearContents
    Range("L15:L305").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
        Range("A14:L30").Select
    ActiveWorkbook.Worksheets("End Of Season Data").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("End Of Season Data").Sort.SortFields.Add Key:= _
        Range("H15:H30"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    ActiveWorkbook.Worksheets("End Of Season Data").Sort.SortFields.Add Key:= _
        Range("J15:J30"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("End Of Season Data").Sort
        .SetRange Range("A14:L30")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    
    
    
    
    Sheets("End Of Season Data").Select
    Range("A15").Select
    Selection.Copy
    Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
    Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
        Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
        Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
        Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
        Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
        Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
        Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
        Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
        Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
        Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
        Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
        Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
        Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
        Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
        Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
            Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
            Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
            Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
            Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
            Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
            Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
            Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
            Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
            Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
            Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
            Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
            Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
            Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
            Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    
            Sheets("Season Groups").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("End Of Season Data").Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -12).Range("A1").Select
    Application.CutCopyMode = False


    Range("L14").Select
    Selection.Copy
    Range("M14").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Group"
    Range("M15").Select
    Columns("M:M").Select
    Selection.NumberFormat = "0"
    Columns("M:M").EntireColumn.AutoFit
    Range("N15").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]>0,1,""*_**_*"")"
    Range("N15").Select
    Selection.Copy
    Range("N16:N100").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=6
    Application.GoTo Reference:="R15C14:R100C14"
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.Replace What:="*_**_*", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Selection.EntireRow.Delete
    Columns("N:N").Select
    Range("N10").Activate
    Selection.ClearContents

    Range("A14").Select
    Application.GoTo Reference:="R14C1:R305C13"
    ActiveWorkbook.Worksheets("End Of Season Data").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("End Of Season Data").Sort.SortFields.Add Key:= _
        Range("M15:M100"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    ActiveWorkbook.Worksheets("End Of Season Data").Sort.SortFields.Add Key:= _
        Range("H15:H100"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    ActiveWorkbook.Worksheets("End Of Season Data").Sort.SortFields.Add Key:= _
        Range("J15:J100"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("End Of Season Data").Sort
        .SetRange Range("A14:M100")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("Q27").Select
    Columns("M:M").Select
    Range("M14").Activate
    Selection.Replace What:="#DIV/0!", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Application.GoTo Reference:="R14C13:R305C13"
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    Range("N1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveCell.FormulaR1C1 = "Highest Winning Percentage"
    Range("N7").Select
    Columns("N:N").EntireColumn.AutoFit
    Range("N7").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.399945066682943
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveCell.FormulaR1C1 = "Most Improved"
    Range("N14").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10092441
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveCell.FormulaR1C1 = "Group Trophies"
    Range("A1").Select
        Range("N15").Select
    ActiveCell.FormulaR1C1 = "=ROUND(RC[-1],0.5)"
    Range("N15").Select
    Selection.Copy
    Application.GoTo Reference:="R15C14:R116C14"
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=-6
    Application.CutCopyMode = False
    Selection.Copy
    Range("M15").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.GoTo Reference:="R15C14:R116C14"
    Application.CutCopyMode = False
    Selection.ClearContents
    Application.GoTo Reference:="R15C13:R116C13"
    Selection.Replace What:="0", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    ActiveWindow.SmallScroll Down:=-200
    Range("A1").Select
    Application.ScreenUpdating = True
End Sub

