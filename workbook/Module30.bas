Attribute VB_Name = "Module30"
Sub Use_LastWeeksLeagueDate()
Attribute Use_LastWeeksLeagueDate.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Use_LastWeeksLeagueDate Macro
'

'
    Application.ScreenUpdating = False
    Range("A18:C18").Select
    Selection.Copy
    Range("A16:C16").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A16:C16").Select
    Range("D17").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C+1"
    Range("D17").Select
    Selection.Copy
    Range("D16").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D17").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("D16").Select
    Application.ScreenUpdating = True
    
End Sub
