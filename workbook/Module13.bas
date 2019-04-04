Attribute VB_Name = "Module13"
Sub Update_Group_Rank()
Attribute Update_Group_Rank.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Updates the group players were in for the end of the season stats
'

'
    Sheets("Season Groups").Select
    Range("D2").Select
    Selection.Copy
    Range("D2:D3000").Select
    ActiveSheet.Paste
    Range("C1").Select
    Selection.ClearContents

    Sheets("Scratch").Select
    Range("A1:ZZ25").Select
    Selection.ClearContents
    Sheets("Groups").Select
    Range("B4:C5").Select
    Application.Run ("Update_Player_hhhhh")
    
    Sheets("Scratch").Select
    Range("A1:ZZ25").Select
    Selection.ClearContents
    Sheets("Groups").Select
    Range("B6:C7").Select
    Application.Run ("Update_Player_hhhhh")
    
    Sheets("Scratch").Select
    Range("A1:ZZ25").Select
    Selection.ClearContents
    Sheets("Groups").Select
    Range("B8:C9").Select
    Application.Run ("Update_Player_hhhhh")
    
    Sheets("Scratch").Select
    Range("A1:ZZ25").Select
    Selection.ClearContents
    Sheets("Groups").Select
    Range("B10:C11").Select
    Application.Run ("Update_Player_hhhhh")

    Sheets("Scratch").Select
    Range("A1:ZZ25").Select
    Selection.ClearContents
    Sheets("Groups").Select
    Range("B12:C13").Select
    Application.Run ("Update_Player_hhhhh")
    
    Sheets("Scratch").Select
    Range("A1:ZZ25").Select
    Selection.ClearContents
    Sheets("Groups").Select
    Range("B14:C15").Select
    Application.Run ("Update_Player_hhhhh")
    
    Sheets("Scratch").Select
    Range("A1:ZZ25").Select
    Selection.ClearContents
    Sheets("Groups").Select
    Range("B16:C17").Select
    Application.Run ("Update_Player_hhhhh")
    
    Sheets("Scratch").Select
    Range("A1:ZZ25").Select
    Selection.ClearContents
    Sheets("Groups").Select
    Range("B18:C19").Select
    Application.Run ("Update_Player_hhhhh")
    
    Sheets("Scratch").Select
    Range("A1:ZZ25").Select
    Selection.ClearContents
    Sheets("Groups").Select
    Range("B20:C21").Select
    Application.Run ("Update_Player_hhhhh")
    

End Sub
