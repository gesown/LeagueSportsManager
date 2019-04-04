VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)


If Target.Address = "$F$13" Then

Call Look_Up_Plyr_By_Rating

End If

If Target.Address = "$E$13" Then

Call Look_Up_Plyr_By_Name

End If

End Sub

