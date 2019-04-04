Attribute VB_Name = "Module38"
Sub FourGroupsEmail()
Attribute FourGroupsEmail.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FourGroupsEmail Macro
'

'
    Range("A12").Select
    ActiveCell.FormulaR1C1 = "=R[-11]C"
    Range("A12").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B12").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Range("A12").Select
End Sub
Sub ThreeGroupsEmail()
Attribute ThreeGroupsEmail.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ThreeGroupsEmail Macro
'

'
    Range("A13").Select
    ActiveCell.FormulaR1C1 = "=R[-11]C"
    Range("A13").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B13").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Range("A13").Select
End Sub
