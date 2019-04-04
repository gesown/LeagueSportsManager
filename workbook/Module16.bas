Attribute VB_Name = "Module16"
Sub selectall()
Attribute selectall.VB_ProcData.VB_Invoke_Func = " \n14"
'
' selectall Macro
'

'
    Range( _
        "F15:G15,I15:J15,L15:M15,O15:P15,O17:P17,L17:M17,I17:J17,F17:G17,F19:G19,I19:J19,L19:M19,O19:P19,O21:P21,L21:M21,I21:J21,F21:G21" _
        ).Select
    Range("F21").Activate
End Sub
