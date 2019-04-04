Attribute VB_Name = "Module23"
Sub MakeRankList()
Attribute MakeRankList.VB_Description = "Creates the rank list for updating on the web site"
Attribute MakeRankList.VB_ProcData.VB_Invoke_Func = " \n14"
'
' MakeRankList Macro
' Creates the rank list for updating on the web site
'

'
    Sheets("Rankings").Select
    Application.GoTo Reference:="R7:R3000"
    Selection.ClearContents
    Sheets("Printable Results").Select
    Application.GoTo Reference:="R7C3:R3000C16"
    Selection.Copy
    Sheets("Rankings").Select
    Range("B7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Printable Results").Select
    Range("S7").Select
    Application.GoTo Reference:="R7C19:R3000C19"
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Rankings").Select
    Range("P7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.GoTo Reference:="R7C4:R3000C4"
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    
    
    Sheets("Printable Results").Select
    Application.GoTo Reference:="R7C16:R3000C16"
    Selection.Copy
    Sheets("Rankings").Select
    Range("N7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    
    Application.Run ("FilterON_ForRankings")
    Range("A1").Select
    Sheets("Printable Results").Select
    Application.GoTo Reference:="R7C4:R3000C4"
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    Application.Run ("FilterON_ForPrintableResults")
    Range("A1").Select
End Sub
Sub DeleteRefError()
'
' DeleteRefError Macro
' Deletes all the #REF! errors in the Home Player List Src before listing the players in the current player list.
'

'
    Sheets("Home Player List Src").Select
    Columns("D:G").Select
    Columns("D:I").Select
    Selection.ClearContents
    Columns("A:L").Select
    Selection.Replace What:="#REF!", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Sheets("Home").Select
End Sub


