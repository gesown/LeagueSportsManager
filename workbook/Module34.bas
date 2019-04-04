Attribute VB_Name = "Module34"
Sub FilterOFF_ForPrintableResults()
Attribute FilterOFF_ForPrintableResults.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FilterOFF_ForPrintableResults Macro
'

'

If ThisWorkbook.Sheets("Printable Results").Range("U1").Value = "Filter" Then

    ActiveSheet.Range("$A$1:$S$3000").AutoFilter Field:=19
    Range("U1").Select
    ActiveCell.FormulaR1C1 = "No Filter"
Else

End If

End Sub
Sub FilterON_ForPrintableResults()
'
' FilterON_ForPrintableResults Macro
'

'

If ThisWorkbook.Sheets("Printable Results").Range("U1").Value = "No Filter" Then

    ActiveSheet.Range("$A$1:$S$3000").AutoFilter Field:=19, Criteria1:="=1", _
        Operator:=xlOr, Criteria2:="="
    Range("U1").Select
    ActiveCell.FormulaR1C1 = "Filter"

Else

End If

End Sub
Sub FilterOFF_ForRankings()
'
' FilterOFF_ForRankings Macro
'

'

If ThisWorkbook.Sheets("Rankings").Range("U1").Value = "Filter" Then

    ActiveSheet.Range("$A$1:$P$3000").AutoFilter Field:=16
    Range("U1").Select
    ActiveCell.FormulaR1C1 = "No Filter"
Else

End If

End Sub
Sub FilterON_ForRankings()
'
' FilterON_ForRankings Macro
'

'

If ThisWorkbook.Sheets("Rankings").Range("U1").Value = "No Filter" Then

    ActiveSheet.Range("$A$1:$P$3000").AutoFilter Field:=16, Criteria1:="=1", _
        Operator:=xlOr, Criteria2:="="
    Range("U1").Select
    ActiveCell.FormulaR1C1 = "Filter"

Else

End If

End Sub
