Attribute VB_Name = "Module7"
Sub Activate_Groups()
Attribute Activate_Groups.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Activate_Groups Macro
'

'


    If MsgBox("Did you check to make sure ALL payers are listed? Please make sure then click yes.", vbYesNo) = vbYes Then
        'continue code
        
    Application.ScreenUpdating = False
    ActiveWorkbook.Save
    Sheets("Search Function").Select
    Range("K1:K3001").Select
    Selection.Copy
    Sheets("Player Archive").Select
    Range("C1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("C2").Select
    Sheets("Attendance").Select
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("C:C").Select
    Selection.Copy
    Columns("D:D").Select
    ActiveSheet.Paste
    Range("D1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Range("E1").Select
    Sheets("Search Function").Select
    Columns("K:K").Select
    Selection.Copy
    Sheets("Attendance").Select
    Columns("D:D").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("D:D").EntireColumn.AutoFit
    Columns("B:B").Select
    Selection.Copy
    Sheets("Player Archive").Select
    Columns("P:P").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Home Player List Src").Select
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC10=1,RC[-4],"""")"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC10=1,RC[-4],"""")"
    Range("E2:F2").Select
    Selection.Copy
    Range("E2:F500").Select
    ActiveSheet.Paste
    Range("E2:F500").Select
    Selection.Copy
    Range("E2:F500").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("E2:F500").Select
    Selection.Copy
    Range("A2:B500").Select
    ActiveSheet.Paste
    Range("E2:F3").Select
    Range("E2:F500").Select
    Selection.ClearContents
    Sheets("Home").Select
    Range("D42").Select
    ActiveCell.FormulaR1C1 = "Ready"
    Application.ScreenUpdating = True
    
    Else
        'exit code
    End If

End Sub
