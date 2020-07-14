Sub 매크로5()
'
' 매크로5 매크로
'

'
    Dim it As Integer, ix As Integer
    
    
    it = ActiveWorkbook.Worksheets.Count    
        
    Cells.Select
    
    For ix = 1 To it
        ActiveWorkbook.Worksheets(ix).Sort.SortFields.Clear
        ActiveWorkbook.Worksheets(ix).Sort.SortFields.Add2 Key:= _
            Range("C2:C500"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        With ActiveWorkbook.Worksheets(ix).Sort
            .SetRange Range("A1:J500")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        ActiveWorkbook.Worksheets(ix).Sort.SortFields.Clear
        ActiveWorkbook.Worksheets(ix).Sort.SortFields.Add2 Key:= _
            Range("C2:C500"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        ActiveWorkbook.Worksheets(ix).Sort.SortFields.Add2 Key:= _
            Range("D2:D500"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortTextAsNumbers
        ActiveWorkbook.Worksheets(ix).Sort.SortFields.Add2 Key:= _
            Range("E2:E500"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        With ActiveWorkbook.Worksheets(ix).Sort
            .SetRange Range("A1:J500")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        Range("A1").Select
    Next ix

End Sub