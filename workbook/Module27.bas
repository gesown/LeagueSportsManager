Attribute VB_Name = "Module27"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Sheets("Season Groups").Select
    Application.GoTo Reference:="R60"
    Rows("60:60").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B59:E59").Select
    Selection.Copy
    Range("B60").Select
    Range("B59:E3176").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
End Sub
