Attribute VB_Name = "Module32"
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'

'
    ActiveSheet.Range("$A$1:$R$182").AutoFilter Field:=16
    ActiveSheet.Range("$A$1:$R$182").AutoFilter Field:=16, Criteria1:=Array("1" _
        , "10", "11", "12", "13", "14", "15", "2", "3", "4", "5", "6", "7", "9", "Total Wks Played", _
        "="), Operator:=xlFilterValues
End Sub
