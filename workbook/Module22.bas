Attribute VB_Name = "Module22"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'

    ActiveWorkbook.Save
    Range("A48:D578").Select
    Selection.Copy
    Sheets("Home Player List Src").Select
    Range("M2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2:I5001").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("M2:O3053").Select
    Selection.Copy
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("P2:P3053").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("M:P").Select
    Application.CutCopyMode = False
    Selection.ClearContents
End Sub
