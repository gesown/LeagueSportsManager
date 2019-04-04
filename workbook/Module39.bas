Attribute VB_Name = "Module39"
Sub Macro9()
Attribute Macro9.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro9 Macro
'

'
    ActiveWorkbook.Worksheets("SeasonWinResults").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SeasonWinResults").Sort.SortFields.Add Key:=Range( _
        "BB4:BB966"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("SeasonWinResults").Sort
        .SetRange Range("AY3:BL966")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
