Attribute VB_Name = "Module12"
Sub Down_Arrow()
Attribute Down_Arrow.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Down_Arrow Macro
'

'
    If ThisWorkbook.Sheets("Home").Range("D42").Value = "Ready" Then
    Application.ScreenUpdating = False
    Sheets("Up Down Arrows").Select
    ActiveWindow.SmallScroll Down:=-15
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[1]:RC[703])"
    Range("A2").Select
    Sheets("Home").Select
    Application.ScreenUpdating = True
    
    Else
    MsgBox ("You must click to start league first before you score the players")
    End If
End Sub
Sub Up_Arrow()
Attribute Up_Arrow.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Up_Arrow Macro
'

'
    If ThisWorkbook.Sheets("Home").Range("D42").Value = "Ready" Then

    Application.ScreenUpdating = False
    Sheets("Up Down Arrows").Select
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[1]:RC[703])"
    Sheets("Home").Select
    Application.ScreenUpdating = True

    Else
    MsgBox ("You must click to start league first before you score the players")
    End If

End Sub
