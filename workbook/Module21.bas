Attribute VB_Name = "Module21"
Sub Go_To_Next_Group()
Attribute Go_To_Next_Group.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Go_To_Next_Group Macro
'

'
    Application.ScreenUpdating = False
    Application.Run ("Detect_And_Update_NBR_of_Players")
    Application.Run ("Update_Group_Rank")
    Sheets("Groups").Select
    Range("A4:A5").Select
    ActiveCell.FormulaR1C1 = "=Home!R[44]C[5]"
    Range("A6:A7").Select
    ActiveCell.FormulaR1C1 = "=Home!R[44]C[5]"
    Range("A8:A9").Select
    ActiveCell.FormulaR1C1 = "=Home!R[44]C[5]"
    Range("A10:A11").Select
    ActiveCell.FormulaR1C1 = "=Home!R[44]C[5]"
    Range("A12:A13").Select
    ActiveCell.FormulaR1C1 = "=Home!R[44]C[5]"
    Range("A14:A15").Select
    ActiveCell.FormulaR1C1 = "=Home!R[44]C[5]"
    Range("A16:A17").Select
    ActiveCell.FormulaR1C1 = "=Home!R[44]C[5]"
    Range("A18:A19").Select
    ActiveCell.FormulaR1C1 = "=Home!R[44]C[5]"
    Range("A20:A21").Select
    ActiveCell.FormulaR1C1 = "=Home!R[44]C[5]"
    Range("A22").Select
    Sheets("Next Group").Select
    Columns("P:P").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("P2").Select

    Sheets("HPLS ").Select
    Range("B2:C295").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("M2").Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   
    Sheets("Left Right Wins").Select
    Selection.ClearContents
    Sheets("Up Down Arrows").Select
    Columns("B:N").Select
    Range("N1").Activate
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[1]:RC[701])"
    Range("A2").Select
    Sheets("Home").Select
    Application.ScreenUpdating = True
End Sub
