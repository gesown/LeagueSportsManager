Attribute VB_Name = "Module28"
Sub PlusFifteen()
'
' PlusFifteen Macro
'

'
    Application.ScreenUpdating = False
    Selection.Copy
    Sheets("plus-min").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]+15"
    Range("B1").Select
    Selection.Copy
    Sheets("Manual Scoring").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("plus-min").Select
    Range("A1:B1").Select
    Range("B1").Activate
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Manual Scoring").Select
    Application.ScreenUpdating = True
End Sub
Sub PlusFive()
'
' PlusFive Macro
'

'
    Application.ScreenUpdating = False
    Selection.Copy
    Sheets("plus-min").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]+5"
    Range("B1").Select
    Selection.Copy
    Sheets("Manual Scoring").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("plus-min").Select
    Range("A1:B1").Select
    Range("B1").Activate
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Manual Scoring").Select
    Application.ScreenUpdating = True
End Sub
Sub MinFifteen()
'
' PlusFifteen Macro
'

'
    Application.ScreenUpdating = False
    Selection.Copy
    Sheets("plus-min").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]-15"
    Range("B1").Select
    Selection.Copy
    Sheets("Manual Scoring").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("plus-min").Select
    Range("A1:B1").Select
    Range("B1").Activate
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Manual Scoring").Select
    Application.ScreenUpdating = True
End Sub
Sub MinFive()
'
' PlusFive Macro
'

'
    Application.ScreenUpdating = False
    Selection.Copy
    Sheets("plus-min").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]-5"
    Range("B1").Select
    Selection.Copy
    Sheets("Manual Scoring").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("plus-min").Select
    Range("A1:B1").Select
    Range("B1").Activate
    Application.CutCopyMode = False
    Selection.ClearContents
    Sheets("Manual Scoring").Select
    Application.ScreenUpdating = True
End Sub

Sub CopyManualRating()
'
' CopyManualRating Macro
'

'
    Application.ScreenUpdating = False
    Calculate
    Range("B21:L22").Select
    Selection.Copy
    Range("B29:L30").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A29:A30").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C29:C30").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C31").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Range("K27").Select
    Application.ScreenUpdating = True
End Sub
Sub CalcManualChange()
'
' CalcManualChange Macro
'

'
    Application.ScreenUpdating = Fales
    Range("E29").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]-RC[-1]"
    Range("E30").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]-RC[-1]"
    Range("K27").Select
    Application.ScreenUpdating = True
End Sub
Sub CalculateSeasonPointsManual()
'
' CalculateSeasonPointsManual Macro
'

'
    Application.ScreenUpdating = False
    Application.Run ("CopyManualRating")
    Range("E29:L30").Select
    Range("L29").Activate
    Selection.Copy
    Sheets("plus-min").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1:A2").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A3").Select
    ActiveSheet.Paste
    Range("H1:H2").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B3:B4").Select
    ActiveSheet.Paste
    Range("C3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]+RC[-2]"
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]+RC[-2]"
    Range("C3:C4").Select
    Selection.Copy
    Range("H1:H2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1:H2").Select
    Range("H1").Activate
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Manual Scoring").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.ScreenUpdating = True
End Sub


