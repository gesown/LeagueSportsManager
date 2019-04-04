Attribute VB_Name = "Module20"
Sub Done_With_League()
Attribute Done_With_League.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Done_With_League Macro
'

'
    Application.ScreenUpdating = False
    
    Sheets("Groups").Select
    Application.Run ("Detect_And_Update_NBR_of_Players")
    Application.Run ("Update_Group_Rank")
    Sheets("Groups").Select
    
    Range("A4:A5").Select
    ActiveCell.FormulaR1C1 = "=Home!R[44]C[5]"
    Range("A6:A7").Select
    ActiveCell.FormulaR1C1 = "=Home!R[44]C[5]"
    Range("A8:A9").Select
    ActiveCell.FormulaR1C1 = "=Home!R[44]C[5]"
    Range("A10:A11").Select
    ActiveCell.FormulaR1C1 = "=Home!R[44]C[5]"
    Range("A12:A13").Select
    ActiveCell.FormulaR1C1 = "=Home!R[44]C[5]"
    Range("A14:A15").Select
    ActiveCell.FormulaR1C1 = "=Home!R[44]C[5]"
    Range("A16:A17").Select
    ActiveCell.FormulaR1C1 = "=Home!R[44]C[5]"
    Range("A18:A19").Select
    ActiveCell.FormulaR1C1 = "=Home!R[44]C[5]"
    Range("A20:A21").Select
    ActiveCell.FormulaR1C1 = "=Home!R[44]C[5]"
    Range("A22").Select
    Sheets("Player Archive").Select
    Cells.Select
    Selection.Copy
    Sheets("Players").Select
    Cells.Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Players").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Players").Sort.SortFields.Add Key:=Range( _
        "E2:E3016"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Players").Sort
        .SetRange Range("A1:U3016")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A1:R3016").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Printable Results").Select
    Range("A6:R6").Select
    ActiveSheet.Paste
    Sheets("Up Down Arrows").Select
    Columns("B:L").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[1]:RC[701])"
    Range("A2").Select
    Sheets("Left Right Wins").Select
    Columns("A:C").Select
    Selection.ClearContents
    Sheets("Update").Select
    Rows("2:2").Select
    Selection.ClearContents
    Sheets("Alphabet Player List").Select
    Columns("AB:AD").Select
    Selection.ClearContents
    Columns("A:C").Select
    Selection.ClearContents
    Sheets("Alpha Names").Select
    Cells.Select
    Selection.ClearContents
    Sheets("Search Function").Select
    Columns("E:H").Select
    Selection.ClearContents
    Columns("M:ALZ").Select
    Selection.ClearContents
    Sheets("Home Player List Src").Select
    Cells.Select
    Selection.ClearContents
    ActiveWindow.ScrollWorkbookTabs Position:=xlFirst
    Sheets("Groups").Select
    Range("O1:ZZ1").Select
    Selection.ClearContents
    Sheets("Next Group").Select
    Range("P1:AZ1").Select
    Selection.ClearContents
    Sheets("Player Archive").Select
    Columns("D:D").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    Sheets("Printable Results").Select
    Application.GoTo Reference:="R7C4:R3300C4"
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    Range("A1:A2").Select
    Sheets("Player Archive").Select
    Range("A1").Select
    Sheets("Home").Select
    Range("D42").Select
    Selection.ClearContents
    Range("G46:H46").Select
    Selection.ClearContents
    Range("S18").Select
    Selection.ClearContents
    Range("S21").Select
    Selection.ClearContents
    ActiveCell.FormulaR1C1 = "Click Start!"
    Range("G26").Select
    Selection.ClearContents
    ActiveCell.FormulaR1C1 = "Ready For League"
    Range("A1").Select
    ActiveWindow.SmallScroll Down:=-93
    Sheets("Season Groups").Select
    Range("D2").Select
    Selection.Copy
    Range("D2:D3000").Select
    ActiveSheet.Paste
    Application.Run ("MakeRankList")
    Application.ScreenUpdating = True
End Sub
