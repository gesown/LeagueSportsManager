Attribute VB_Name = "Module25"
Sub Macro6()
Attribute Macro6.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro6 Macro
'

'

    ActiveWindow.ScrollWorkbookTabs Position:=xlFirst
    Sheets("Groups").Select
    Range("B4:C5").Select
    Application.CutCopyMode = False
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
    ActiveCell.Offset(0, 4).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("Groups").Select
    ActiveCell.Offset(-2, -1).Range("A1").Select
    Selection.Copy
    Sheets("Season Groups").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    Sheets("Groups").Select
    ActiveWindow.SmallScroll Down:=21
    Range("N38").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Range("N39").Select
    ActiveWindow.SmallScroll Down:=-33
End Sub
