Attribute VB_Name = "Module5"

Sub Alphabetize_Players()
Attribute Alphabetize_Players.VB_Description = "Alphabetizes All The Players So You Can Search And Add Them To The League"
Attribute Alphabetize_Players.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Alphabetize_Players Macro
' Alphabetizes All The Players So You Can Search And Add Them To The League
'

'

    If MsgBox("Are you ready to start league? You MUST add new players before you start!", vbYesNo) = vbYes Then
        'continue code

    Application.ScreenUpdating = False
    Sheets("SeasonWinResults").EnableCalculation = False
    
    Sheets("Printable Results").Select
    Application.Run ("FilterOFF_ForPrintableResults")
    Sheets("Rankings").Select
    Application.Run ("FilterOFF_ForRankings")
    
    Sheets("Players").Select
    Columns("A:S").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Player Archive").Select
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Columns("A:S").Select
    Selection.Copy
    Sheets("Players").Select
    Columns("A:S").Select
    ActiveSheet.Paste
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Columns("D:E").Select
    Selection.Copy
    Sheets("Alpha Names").Select
    Columns("A:H").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Columns("A:B").Select
    Sheets("Players").Select
    Columns("D:E").Select
    Selection.Copy
    Sheets("Alpha Names").Select
    Columns("A:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Rows("1:1").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Columns("A:B").Select
    ActiveWorkbook.Worksheets("Alpha Names").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Alpha Names").Sort.SortFields.Add Key:=Range( _
        "A1:A1002"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Alpha Names").Sort
        .SetRange Range("A1:B1002")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.Copy
    Sheets("Alphabet Player List").Select
    Columns("AB:AC").Select
    ActiveSheet.Paste
    Columns("AB:AB").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Alpha Names").Select
    Columns("D:D").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("D1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 2), Array(2, 1)), TrailingMinusNumbers:=True
    Columns("D:F").Select
    Selection.Copy
    Sheets("Alphabet Player List").Select
    Columns("A:C").Select
    ActiveSheet.Paste
    Sheets("Search Function").Select
    Columns("E:H").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Search Function").Select
    Columns("M:ALX").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Home").Select
    Range("G26:J26").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    ActiveCell.FormulaR1C1 = "Players Are Now Alphabetized"
    Application.ScreenUpdating = True
    
        Else
    Range("F16").Select
    End If
End Sub
