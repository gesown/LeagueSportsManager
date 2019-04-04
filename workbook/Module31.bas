Attribute VB_Name = "Module31"
Sub Detect_And_Update_NBR_of_Players()
Attribute Detect_And_Update_NBR_of_Players.VB_Description = "finds the sum of Groups tab column A4 - A21 and updates the number of players according to how many people there were in the group."
Attribute Detect_And_Update_NBR_of_Players.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Detect_And_Update_NBR_of_Players Macro
' finds the sum of Groups tab column A4 - A21 and updates the number of players according to how many people there were in the group.
'

'

Application.ScreenUpdating = False
s = ActiveWorkbook.Worksheets("Groups").Range("B1")
Application.ScreenUpdating = False

If s = 15 Then GoTo 15:
If s = 21 Then GoTo 21:
If s = 28 Then GoTo 28:
If s = 36 Then GoTo 36:
If s = 45 Then GoTo 45:

MsgBox ("Sum is out of expected range of possible # of players")
Exit Sub

15:

Call Update_Player_One
Call Update_Player_Two
Call Update_Player_Three
Call Update_Player_Four
Call Update_Player_Five

Exit Sub

21:

Call Update_Player_One
Call Update_Player_Two
Call Update_Player_Three
Call Update_Player_Four
Call Update_Player_Five
Call Update_Player_Six

Exit Sub

28:

Call Update_Player_One
Call Update_Player_Two
Call Update_Player_Three
Call Update_Player_Four
Call Update_Player_Five
Call Update_Player_Six
Call Update_Player_Seven

Exit Sub

36:

Call Update_Player_One
Call Update_Player_Two
Call Update_Player_Three
Call Update_Player_Four
Call Update_Player_Five
Call Update_Player_Six
Call Update_Player_Seven
Call Update_Player_Eight

Exit Sub

45:

Call Update_Player_One
Call Update_Player_Two
Call Update_Player_Three
Call Update_Player_Four
Call Update_Player_Five
Call Update_Player_Six
Call Update_Player_Seven
Call Update_Player_Eight
Call Update_Player_Nine

Exit Sub


End Sub
