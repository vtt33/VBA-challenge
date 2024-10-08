Sub quarterly_change()
    ' Create variables
    Dim i As Long
    Dim ws As Worksheet
    Dim LastRow As Long
    
    ' Create variables to hold Ticker, openingPrice, closingPrice, quarterChange, stockDate, currentQuarter, and totalVolume
    Dim Ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim quarterChange As Double
    Dim currentDate As Date
    Dim totalVolume As Double
    
    ' Initialize for the values
    openingPrice = 0
    closingPrice = 0
    quarterChange = 0
    totalVolume = 0
    
    ' Variables for tracking greatest values
    Dim maxPercentChange As Double
    Dim minPercentChange As Double
    Dim maxVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    
    
    ' Initialize greatest values
    maxPercentChange = -1
    minPercentChange = 1
    maxVolume = 0
    
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
        
        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Add Columns for output
        ws.Range("I1").EntireColumn.Insert
        ws.Range("J1").EntireColumn.Insert
        ws.Range("K1").EntireColumn.Insert
        ws.Range("L1").EntireColumn.Insert
        
        ' Label Output Columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Keep track of the location for outputs in the summary table
        Dim Summary_Row As Long
        Summary_Row = 2
        
        ' Loop through rows and store values in appropriate columns
        For i = 2 To LastRow
            
    
            ' Set the ticker
            Ticker = ws.Cells(i, 1).Value
            
            ' Variables for previous row values
            Dim previousTicker As String
            Dim previousDate As Date
            
            If i > 2 Then ' Only check previous values if not the first row
                previousTicker = ws.Cells(i - 1, 1).Value
                previousDate = ws.Cells(i - 1, 2).Value
                
            End If
            
            ' Determine current quarter
            currentDate = ws.Cells(i, 2).Value
            
            ' If it's the start of a new quarter or a new ticker
            If i = 2 Or previousTicker <> Ticker Or (previousDate <> 0 And ws.Cells(i, 2).Value <> currentDate) Then
                
                ' Set the opening price
                openingPrice = ws.Cells(i, 3).Value
                totalVolume = 0 ' Reset total volume for the new quarter
                
            End If
            
            ' Accumulate the closing price and total volume
            closingPrice = ws.Cells(i, 6).Value
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' If it's the end of the last entry for this ticker
            If i = LastRow Or ws.Cells(i + 1, 1).Value <> Ticker Then
                ' Calculate quarterly change
                quarterChange = closingPrice - openingPrice
                
                ' Calculate percentage change
                Dim percentChange As Double
                If openingPrice <> 0 Then
                    percentChange = (quarterChange / openingPrice)
                Else
                    percentChange = 0
                End If
                
                ' Print the Tickers in the Summary Table
                ws.Range("I" & Summary_Row).Value = Ticker
                
                ' Print the quarterChange to the Summary Table and format the color
                ws.Range("J" & Summary_Row).Value = quarterChange
                
                If quarterChange > 0 Then
                    ws.Range("J" & Summary_Row).Interior.ColorIndex = 4
                    
                    ElseIf (quarterChange < 0) Then
                    
                        ws.Range("J" & Summary_Row).Interior.ColorIndex = 3
                    
                    
                End If
                    
             
                ' Print the percentChange to the Summary Table
                ws.Range("K" & Summary_Row).Value = percentChange
                ws.Range("K" & Summary_Row).NumberFormat = "0.00%"
                
                ' Print the total volume to the Summary Table
                ws.Range("L" & Summary_Row).Value = totalVolume
                ws.Range("L" & Summary_Row).NumberFormat = "0"
                
                ' Increment the summary row
                Summary_Row = Summary_Row + 1
                
                ' Check for greatest % increase, greatest % decrease and total volume
                If percentChange > maxPercentChange Then
                    maxPercentChange = percentChange
                    greatestIncreaseTicker = Ticker
                    
                End If
                
                If percentChange < minPercentChange Then
                    minPercentChange = percentChange
                    greatestDecreaseTicker = Ticker
                    
                End If
                
                If totalVolume > maxVolume Then
                    maxVolume = totalVolume
                    greatestVolumeTicker = Ticker
                    
                End If
                
                ' Reset values for next ticker/quarter
                openingPrice = 0
                closingPrice = 0
                quarterChange = 0
                totalVolume = 0
            End If
            
        Next i
        
        ' Label for the header column
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ' Output the results for Greatest % Increase, % Decrease, and Total Volume
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = greatestIncreaseTicker
        ws.Cells(2, 17).Value = maxPercentChange
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        
        ' Output the results for Greatest % Decrease
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = greatestDecreaseTicker
        ws.Cells(3, 17).Value = minPercentChange
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        
        
        ' Output the results for Greatest Total Volume
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = greatestVolumeTicker
        ws.Cells(4, 17).Value = maxVolume
        
        
    ' --------------------------------------------
    ' TASKS COMPLETE
    ' --------------------------------------------
    Next ws

    MsgBox "Tasks Complete"
End Sub


