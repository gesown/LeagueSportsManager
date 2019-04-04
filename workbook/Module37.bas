Attribute VB_Name = "Module37"
Sub ScrollUpPlyr()
Attribute ScrollUpPlyr.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ScrollUpPlyr Macro
'

'
If Sheets("Home").Range("T26").Value = 1 Then

Sheets("Home").Range("U26").Value = 1

Else
CurrentLineNBR = Sheets("Home").Range("U26").Value
Sheets("Home").Range("U26").Value = CurrentLineNBR - 1

End If

End Sub
Sub ScrollDwnPlyr()
'
' ScrollUpPlyr Macro
'

'

CurrentLineNBR = Sheets("Home").Range("U26").Value
Sheets("Home").Range("U26").Value = CurrentLineNBR + 1

End Sub
Sub CenterIt()
Attribute CenterIt.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CenterIt Macro
'

'
    Sheets("Home").Range("U26").Value = 0
End Sub
