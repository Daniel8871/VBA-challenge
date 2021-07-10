Sub Market()

dim CurrentTicker as String, NextTicker as String
dim TotalVolume as LongLong
dim RowNum as Long, RowNum_open as Long, SummaryRow as Integer, LastRow as Long
dim firstOpen as Integer, PChange as Integer

For each ws in Worksheets 
    SummaryRow = 2
    RowNum_open = 2
    LastRow=ws.Cells(rows.count,1).End(xlUp).Row

    ws.cells(1,9).value="Ticker"
    ws.cells(1,10).value="Yearly Change"
    ws.cells(1,11).value="Percent Change"
    ws.cells(1,12).value="Total Stock Volume"

    firstOpen=ws.cells(RowNum_open,3).value
    For RowNum = 2 to LastRow
        CurrentTicker= ws.cells(RowNum,1).value
        NextTicker=ws.cells(RowNum+1,1).value
        TotalVolume=TotalVolume+ws.cells(RowNum,7).value
        YChange=ws.cells(RowNum,6).value-firstOpen
        PChange=(ws.cells(RowNum,6)-firstOpen)/firstOpen

        If CurrentTicker <> NextTicker Then
            ws.cells(SummaryRow,9).value=CurrentTicker
            ws.cells(SummaryRow,10).value=YChange
            ws.cells(SummaryRow,11).value=PChange
            ws.cells(SummaryRow,12).value=TotalVolume
            
            RowNum_open=RowNum+1
            SummaryRow=SummaryRow+1
            TotalVolume=0
        end if        
    Next RowNum 
    ws.range("K1:K"&LastRow).NumberFormat="0%"
next ws 

End Sub