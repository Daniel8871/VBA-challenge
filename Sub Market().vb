Sub Market()

dim CurrentTicker as String, NextTicker as String
dim TotalVolume as LongLong, Vol_Sum as LongLong, P_Sum as Double
dim RowNum as Long, RowNum_open as Long, SummaryRow as Integer, LastRow as Long
dim firstOpen as Double, PChange as Double, YChange as Double

For each ws in Worksheets 

    SummaryRow = 2 'Iterable row used in summary columns 9-12
    RowNum_open = 2 'Iterable row used to search through all rows of data for <open> price
    LastRow=ws.Cells(rows.count,1).End(xlUp).Row 

    'I'm sure there is a cleaner way to write this, 
    'but this is my header line of summary table
    ws.cells(1,9).value="Ticker"
    ws.cells(1,10).value="Yearly Change"
    ws.cells(1,11).value="Percent Change"
    ws.cells(1,12).value="Total Stock Volume"

    firstOpen=ws.cells(RowNum_open,3).value
    For RowNum = 2 to LastRow
        
        'Finding ticker transition
        CurrentTicker= ws.cells(RowNum,1).value
        NextTicker=ws.cells(RowNum+1,1).value

        'Statistical calculations
        TotalVolume=ws.cells(RowNum,7).value+TotalVolume
        YChange=ws.cells(RowNum,6).value-firstOpen
        PChange=YChange/firstOpen

        If CurrentTicker <> NextTicker Then

            'Fill in summary table for completed ticker data set
            ws.cells(SummaryRow,9).value=CurrentTicker
            ws.cells(SummaryRow,10).value=YChange
                'format
                If YChange > 0 Then 
                    ws.cells(SummaryRow,10).interior.colorIndex=4 'green
                Else ws.cells(SummaryRow,10).interior.colorIndex=3 'red
                End if
            ws.cells(SummaryRow,11).value=PChange
                'format
                ws.cells(SummaryRow,11).NumberFormat="0.00%"
            ws.cells(SummaryRow,12).value=TotalVolume
                       
            'first row of new ticker is the <open> price
            RowNum_open=RowNum+1
            firstOpen=ws.cells(RowNum_open,3).value

            'prepare summary row
            SummaryRow=SummaryRow+1

            'reset Total Volume for ticker
            TotalVolume=0
        end if        

    Next RowNum 
    
    Set Summary_range = ws.range("K2:K"&SummaryRow)
        ws.cells(1,15).value="Greatest % Decrease:"
        P_Sum=Application.WorksheetFunction.Min(Summary_range)
        ws.cells(1,16).value=P_Sum
            ws.cells(1,16).NumberFormat="0.00%"

        ws.cells(2,15).value="Greatest % Increase:"
        P_Sum=Application.WorksheetFunction.Max(Summary_range)
        ws.cells(2,16).value=P_Sum
            ws.cells(2,16).NumberFormat="0.00%"

    Set Vol_Summary_range = ws.range("L2:L"&SummaryRow)
        ws.cells(3,15).value="Greatest Volume:"
        Vol_Sum=Application.WorksheetFunction.Max(Vol_Summary_range)
            ws.cells(3,16).value=Vol_Sum
    Vol_Sum=0
    P_Sum=0
next ws 

End Sub