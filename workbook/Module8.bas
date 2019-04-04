Attribute VB_Name = "Module8"
Sub Custom_Player_Order()
Attribute Custom_Player_Order.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Custom_Player_Order Macro
'

'
    If ThisWorkbook.Sheets("Home").Range("D42").Value = "Ready" Then
    ActiveWindow.ScrollWorkbookTabs Position:=xlLast
    Sheets("Custom Order").Select
    Range("F9").Select
    Else
    MsgBox ("You must click to start league first before you set the player order")
    End If

End Sub
Sub Reorganize_Player_Order()
Attribute Reorganize_Player_Order.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Reorganize_Player_Order Macro
'

'
    Application.ScreenUpdating = False
    Range("A8:J493").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("D:D").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("B:B").Select
    ActiveSheet.Paste
    Columns("A:J").Select
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Home Player List Src").Sort.SortFields.Add Key:= _
        Range("F2:F3001"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Home Player List Src").Sort
        .SetRange Range("A1:J3001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Custom Order").Select
    Range("F9:F493").Select
    Selection.ClearContents
    Sheets("Home").Select
    Range("H46").Select
    ActiveCell.FormulaR1C1 = "Done!"
    Range("G46").Select
    ActiveCell.FormulaR1C1 = "Re-order"
    Application.ScreenUpdating = True
End Sub
