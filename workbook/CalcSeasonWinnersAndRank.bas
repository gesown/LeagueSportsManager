Attribute VB_Name = "CalcSeasonWinnersAndRank"
Sub CalculteSeasonWinners()
'
'
'
    Sheets("SeasonWinResults").Range("AY3:BL1000").Value = Sheets("SeasonWinResults").Range("S3:AF1000").Value
    Sheets("SeasonWinResults").Range("AY3:BL966").Select
    ActiveWorkbook.Worksheets("SeasonWinResults").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SeasonWinResults").Sort.SortFields.Add Key:=Range( _
        "AZ4:AZ966"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("SeasonWinResults").Sort.SortFields.Add Key:=Range( _
        "BH4:BH966"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("SeasonWinResults").Sort.SortFields.Add Key:=Range( _
        "BJ4:BJ966"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("SeasonWinResults").Sort.SortFields.Add Key:=Range( _
        "BL4:BL966"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
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
Sub CalculteWeeksRanks()
'
'
'
    Sheets("SeasonWinResults").EnableCalculation = True
    Sheets("SeasonWinResults").Select
    Sheets("SeasonWinResults").Range("AY3:BL1000").Value = Sheets("SeasonWinResults").Range("S3:AF1000").Value
    Sheets("SeasonWinResults").Range("AY3:BL3000").Select
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
    Sheets("SeasonWinResults").EnableCalculation = False
    Sheets("SeasonWinResults").Range("CQ1").Select
End Sub
Sub CopyWeeksRankings()
'
'
'
CopyRankingsArea = Sheets("SeasonWinResults").Range("CE1").Value
Sheets("SeasonWinResults").Range(CopyRankingsArea).Copy

End Sub
Sub CalcsManual()
'
'
'

Sheets("SeasonWinResults").EnableCalculation = True

Sheets("SeasonWinResults").EnableCalculation = False

End Sub
