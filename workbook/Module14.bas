Attribute VB_Name = "Module14"
Sub Left_Player_Wins()
Attribute Left_Player_Wins.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Left_Player_Wins Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Home").Select
    Range("Z46").Select
    Selection.Copy
    Sheets("Left Right Wins").Select
    Range("B25").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A25").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "1"
    Columns("A:B").Select
    ActiveWorkbook.Worksheets("Left Right Wins").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Left Right Wins").Sort.SortFields.Add Key:=Range( _
        "B1:B25"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Left Right Wins").Sort
        .SetRange Range("A1:B25")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("D19").Select
    ActiveCell.FormulaR1C1 = "a"
    Range("D18").Select
    ActiveCell.FormulaR1C1 = "b"
    Range("D17").Select
    ActiveCell.FormulaR1C1 = "c"
    Range("D16").Select
    ActiveCell.FormulaR1C1 = "d"
    Range("D15").Select
    ActiveCell.FormulaR1C1 = "e"
    Range("D14").Select
    ActiveCell.FormulaR1C1 = "f"
    Range("D13").Select
    ActiveCell.FormulaR1C1 = "g"
    Range("D12").Select
    ActiveCell.FormulaR1C1 = "h"
    Range("D11").Select
    ActiveCell.FormulaR1C1 = "i"
    Range("D10").Select
    ActiveCell.FormulaR1C1 = "j"
    Range("D9").Select
    ActiveCell.FormulaR1C1 = "k"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "l"
    Range("D7").Select
    ActiveCell.FormulaR1C1 = "m"
    Range("D6").Select
    ActiveCell.FormulaR1C1 = "n"
    Range("D5").Select
    ActiveCell.FormulaR1C1 = "o"
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "p"
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "q"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "r"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "s"
    Columns("A:D").Select
    Range("D1").Activate
    ActiveWorkbook.Worksheets("Left Right Wins").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Left Right Wins").Sort.SortFields.Add Key:=Range( _
        "D1:D25"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Left Right Wins").Sort
        .SetRange Range("A1:D25")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("A:B").Select
    Range("B1").Activate
    ActiveSheet.Range("$A$1:$B$25").RemoveDuplicates Columns:=2, Header:=xlNo
    Columns("D:D").Select
    Selection.ClearContents
    Columns("A:B").Select
    ActiveWorkbook.Worksheets("Left Right Wins").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Left Right Wins").Sort.SortFields.Add Key:=Range( _
        "B1:B25"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Left Right Wins").Sort
        .SetRange Range("A1:B25")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
        Sheets("Home").Select
    Sheets("Up Down Arrows").Select
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[1]:RC[703])"
    Range("A2").Select
    Sheets("Home").Select
    Application.ScreenUpdating = True
End Sub
Sub Right_Player_Wins()
Attribute Right_Player_Wins.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Right_Player_Wins Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Home").Select
    Range("Z46").Select
    Range("Z46").Select
    Selection.Copy
    Sheets("Left Right Wins").Select
    Range("B25").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A25").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "2"
    Columns("A:B").Select
    ActiveWorkbook.Worksheets("Left Right Wins").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Left Right Wins").Sort.SortFields.Add Key:=Range( _
        "B1:B25"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Left Right Wins").Sort
        .SetRange Range("A1:B25")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("D20").Select
    ActiveCell.FormulaR1C1 = "a"
    Range("D19").Select
    ActiveCell.FormulaR1C1 = "b"
    Range("D18").Select
    ActiveCell.FormulaR1C1 = "c"
    Range("D17").Select
    ActiveCell.FormulaR1C1 = "d"
    Range("D16").Select
    ActiveCell.FormulaR1C1 = "e"
    Range("D15").Select
    ActiveCell.FormulaR1C1 = "f"
    Range("D14").Select
    ActiveCell.FormulaR1C1 = "g"
    Range("D13").Select
    ActiveCell.FormulaR1C1 = "h"
    Range("D12").Select
    ActiveCell.FormulaR1C1 = "i"
    Range("D11").Select
    ActiveCell.FormulaR1C1 = "j"
    Range("D10").Select
    ActiveCell.FormulaR1C1 = "k"
    Range("D9").Select
    ActiveCell.FormulaR1C1 = "l"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "m"
    Range("D7").Select
    ActiveCell.FormulaR1C1 = "n"
    Range("D6").Select
    ActiveCell.FormulaR1C1 = "o"
    Range("D5").Select
    ActiveCell.FormulaR1C1 = "p"
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "q"
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "r"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "s"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "t"
    Columns("A:D").Select
    Range("D1").Activate
    ActiveWorkbook.Worksheets("Left Right Wins").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Left Right Wins").Sort.SortFields.Add Key:=Range( _
        "D1:D25"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Left Right Wins").Sort
        .SetRange Range("A1:D25")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("A:B").Select
    Range("B10").Activate
    ActiveSheet.Range("$A$1:$B$25").RemoveDuplicates Columns:=2, Header:=xlNo

    Columns("D:D").Select
    Selection.ClearContents
    Columns("A:B").Select
    Range("B1").Activate
    ActiveWorkbook.Worksheets("Left Right Wins").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Left Right Wins").Sort.SortFields.Add Key:=Range( _
        "B1:B25"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Left Right Wins").Sort
        .SetRange Range("A1:B25")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
        Sheets("Home").Select
    Sheets("Up Down Arrows").Select
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[1]:RC[703])"
    Range("A2").Select
    Sheets("Home").Select
    Application.ScreenUpdating = True
End Sub
Sub No_Play()
Attribute No_Play.VB_ProcData.VB_Invoke_Func = " \n14"
'
' No_Play Macro
'

'
    Application.ScreenUpdating = False
    Range("Z46").Select
    Selection.Copy
    Sheets("Left Right Wins").Select
    Range("B25").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A25").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "0"
    Columns("A:B").Select
    Range("B1").Activate
    ActiveWorkbook.Worksheets("Left Right Wins").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Left Right Wins").Sort.SortFields.Add Key:=Range( _
        "B1:B25"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Left Right Wins").Sort
        .SetRange Range("A1:B25")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("D20").Select
    ActiveCell.FormulaR1C1 = "a"
    Range("D19").Select
    ActiveCell.FormulaR1C1 = "b"
    Range("D18").Select
    ActiveCell.FormulaR1C1 = "c"
    Range("D17").Select
    ActiveCell.FormulaR1C1 = "d"
    Range("D16").Select
    ActiveCell.FormulaR1C1 = "e"
    Range("D15").Select
    ActiveCell.FormulaR1C1 = "f"
    Range("D14").Select
    ActiveCell.FormulaR1C1 = "g"
    Range("D13").Select
    ActiveCell.FormulaR1C1 = "h"
    Range("D12").Select
    ActiveCell.FormulaR1C1 = "i"
    Range("D11").Select
    ActiveCell.FormulaR1C1 = "j"
    Range("D10").Select
    ActiveCell.FormulaR1C1 = "k"
    Range("D9").Select
    ActiveCell.FormulaR1C1 = "l"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "m"
    Range("D7").Select
    ActiveCell.FormulaR1C1 = "n"
    Range("D6").Select
    ActiveCell.FormulaR1C1 = "o"
    Range("D5").Select
    ActiveCell.FormulaR1C1 = "p"
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "q"
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "r"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "s"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "t"
    Columns("A:D").Select
    Range("D1").Activate
    ActiveWorkbook.Worksheets("Left Right Wins").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Left Right Wins").Sort.SortFields.Add Key:=Range( _
        "D1:D25"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Left Right Wins").Sort
        .SetRange Range("A1:D25")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("A:B").Select
    Range("B1").Activate
    ActiveSheet.Range("$A$1:$B$25").RemoveDuplicates Columns:=2, Header:=xlNo
    Columns("D:D").Select
    Selection.ClearContents
    Columns("A:B").Select
    Range("B1").Activate
    ActiveWorkbook.Worksheets("Left Right Wins").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Left Right Wins").Sort.SortFields.Add Key:=Range( _
        "B1:B25"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Left Right Wins").Sort
        .SetRange Range("A1:B25")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Home").Select
    Sheets("Up Down Arrows").Select
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[1]:RC[703])"
    Range("A2").Select
    Sheets("Home").Select
    Application.ScreenUpdating = True
End Sub
Sub Done_Scoring_Group()
Attribute Done_Scoring_Group.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Done_Scoring_Group Macro
'

'

    If ThisWorkbook.Sheets("Home").Range("D42").Value = "Ready" Then
    Sheets("Groups").Select
    ActiveWorkbook.Save
    Range("A27").Select
    ActiveWindow.SmallScroll Down:=-15
    
    Else
    MsgBox ("You must click to start league first before you score the players")
    End If
    
End Sub
