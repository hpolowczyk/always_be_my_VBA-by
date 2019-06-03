Sub NotSoHardVBA()

    '------------------------------------------------------
    ' Looping through all sheets
    '------------------------------------------------------
    For Each ws In Worksheets

        ' Determine the last row and column of the provided data
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        lastcolumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        '--------------------------------------------------
        ' Find all unique ticker symbols
        '--------------------------------------------------

        ' Set the header
        ws.Range("I1").Value = "Ticker"

        ' Copy over <ticker> (i.e. Column A) into Ticker (i.e. Column I)
        ws.Range("A2:A" & lastrow).Copy Destination:=ws.Range("I2:I" & lastrow)

        ' Remove all duplicates in Ticker 
        ws.Range("I2:I" & lastrow).RemoveDuplicates Columns:=Array(1)

        '--------------------------------------------------
        ' Determine:
        ' (1) Yearly price change 
        ' (2) Percent price change 
        ' (3) Total stock volume
        '--------------------------------------------------

        ' Determine the last row of Ticker
        tickerrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

        ' Set the headers
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        ' Define and set ranges
        Dim TickerTotal As Range
        Dim Dates As Range
        Dim OpenPrice As Range
        Dim ClosePrice As Range
        Dim SumTotal As Range
        
        Set TickerTotal = ws.Range("A2:A" & lastrow)
        Set Dates = ws.Range("B2:B" & lastrow)
        Set OpenPrice = ws.Range("C2:C" & lastrow)
        Set ClosePrice = ws.Range("F2:F" & lastrow)
        Set SumTotal = ws.Range("G2:G" & lastrow)
        
        '--------------------------------------------------
        ' (1) YEARLY PRICE CHANGE
        '--------------------------------------------------

        ' Find the difference between initial open price and final close price by looping through each stock
        For i = 2 To tickerrow
           
            ' Set variables 
            Dim diff As Double
            Dim initial_open_price As Double
            Dim final_close_price As Double
            
            ' Use MINIFS to find first date of the given year for each stock
            initial_date = Application.WorksheetFunction.MinIfs(Dates, TickerTotal, ws.Cells(i, 9).Value)
            ' Use SUMIFS to find open price under the condition that the date matches initial_date and the ticker matches the stock
            initial_open_price = Application.WorksheetFunction.SumIfs(OpenPrice, TickerTotal, ws.Cells(i, 9).Value, Dates, initial_date)
            
            ' Use MAXIFS to find last date of the given year for each stock
            final_date = Application.WorksheetFunction.MaxIfs(Dates, TickerTotal, ws.Cells(i, 9).Value)
            ' Use SUMIFS to find close price under the condition that the date matches final_date and the ticker matches the stock
            final_close_price = Application.WorksheetFunction.SumIfs(ClosePrice, TickerTotal, ws.Cells(i, 9).Value, Dates, final_date)
            
            ' Find the difference between final_close_price and initial_open_price
            diff = final_close_price - initial_open_price

            ' Set diff to each cell value in the Yearly Change column (i.e. Column J)
            ws.Cells(i, 10).Value = diff
            
            ' Set cell colour to green if diff is greater than or equal zero, if not...
            If ws.Cells(i, 10).Value >= 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            ' Set cell colour to red if diff is less than zero
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
            
            '----------------------------------------------
            ' (2) PERCENT PRICE CHANGE
            '----------------------------------------------
            
            ' Since division by 0 results in an error, if initial_open_price is zero, then set percent change to zero
            If initial_open_price = 0 Then
                ws.Cells(i, 11).Value = 0
            ' If initial_open_price is not zero, divide diff by initial_open_price to get the yearly percent price change
            Else
                ws.Cells(i, 11).Value = diff / initial_open_price
            End If
            
            ' Set cell format to percentage with 2 decimal places
            ws.Cells(i, 11).NumberFormat = "0.00%"
            
            '----------------------------------------------
            ' (3) TOTAL STOCK VOLUME
            '----------------------------------------------
            
            ' Use SUMIFS to add all stock volumes given that the ticker matches the stock
            ws.Cells(i, 12).Value = Application.WorksheetFunction.SumIfs(SumTotal, TickerTotal, ws.Cells(i, 9).Value)

        Next i
        
        ' -------------------------------------------------
        ' Data analysis
        ' -------------------------------------------------
        
        ' Set the headers
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ' Define and set ranges
        Dim PercentChange As Range
        Dim TotalVolume As Range
        
        Set PercentChange = ws.Range("K2:K" & tickerrow)
        Set TotalVolume = ws.Range("L2:L" & lastrow)
        
        '--------------------------------------------------
        ' (1) GREATEST % INCREASE
        '--------------------------------------------------

        ' Set row title
        ws.Range("O2").Value = "Greatest % Increase"
        ' Use MAX to determine the largest percent change 
        great_incr = Application.WorksheetFunction.Max(PercentChange)
        ' Use MATCH to identify the row location of great_inc
        great_incr_row = Application.WorksheetFunction.Match(great_incr, PercentChange, 0) + 1
        ' Identity the stock with the greatest percent increase
        ws.Cells(2, 16).Value = ws.Cells(great_incr_row, 9).Value
        ' Identity the percent change
        ws.Cells(2, 17).Value = ws.Cells(great_incr_row, 11).Value
        ' Set cell format to percentage with 2 decimal places
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        '--------------------------------------------------
        ' (2) GREATEST % DECREASE
        '--------------------------------------------------

        ' Set row title
        ws.Range("O3").Value = "Greatest % Decrease"
        ' Use MIN to determine the smallest percent change
        great_decr = Application.WorksheetFunction.Min(PercentChange)
        ' Use MATCH to identify the row location of great_decr
        great_decr_row = Application.WorksheetFunction.Match(great_decr, PercentChange, 0) + 1
        ' Identity the stock with the greatest percent decrease
        ws.Cells(3, 16).Value = ws.Cells(great_decr_row, 9).Value
        ' Identity the percent change
        ws.Cells(3, 17).Value = ws.Cells(great_decr_row, 11).Value
        ' Set cell format to percentage with 2 decimal places
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        '--------------------------------------------------
        ' (3) GREATEST TOTAL VOLUME
        '--------------------------------------------------

        ' Set row title
        ws.Range("O4").Value = "Greatest Total Volume"
        ' Use MAX to determine the largest total volume
        great_vol = Application.WorksheetFunction.Max(TotalVolume)
        ' Use MATCH to identify the row location of great_inc
        great_vol_row = Application.WorksheetFunction.Match(great_vol, TotalVolume, 0) + 1
        ' Identity the stock with the greatest total volume
        ws.Cells(4, 16).Value = ws.Cells(great_vol_row, 9).Value
        ' Identity the total volume amount
        ws.Cells(4, 17).Value = ws.Cells(great_vol_row, 12).Value


        ' Adjust columns to ensure all the data fits
        ws.Columns("I:Q").AutoFit

    Next ws

End Sub
