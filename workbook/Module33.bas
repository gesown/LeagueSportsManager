Attribute VB_Name = "Module33"
Sub Macro5()
Attribute Macro5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro5 Macro
'

'
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
End Sub
