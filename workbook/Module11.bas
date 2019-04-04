Attribute VB_Name = "Module11"
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'


    Selection.Copy
    ActiveWindow.ScrollWorkbookTabs Position:=xlLast
    Sheets("Scratch").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveWindow.ScrollWorkbookTabs Position:=xlFirst
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
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Rows("1:1").EntireRow.Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveWindow.ScrollWorkbookTabs Position:=xlLast
    Sheets("Scratch").Select
    ActiveCell.Rows("1:1").EntireRow.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, 5).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollWorkbookTabs Position:=xlFirst
    Sheets("Groups").Select
    Range("A2").Select
    Selection.Copy
    ActiveWindow.ScrollWorkbookTabs Position:=xlLast
    Sheets("Scratch").Select
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Range("G2").Select
    Selection.Copy
    Range("G1").Select
    ActiveSheet.Paste
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("G2").Select
    Selection.ClearContents
    Rows("1:1").Select
    Selection.Copy
    ActiveWindow.ScrollWorkbookTabs Position:=xlFirst
    Sheets("Season Groups").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").Select
    Selection.Find(What:="50", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Offset(1, 1).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveCell.Offset(-1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 1).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveCell.Offset(-1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 1).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveCell.Offset(-1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.Select
    Sheets("Groups").Select
    ActiveWindow.SmallScroll Down:=18
    ActiveCell.Offset(38, 16).Range("A1").Select
    ActiveWindow.SmallScroll Down:=-39
End Sub

