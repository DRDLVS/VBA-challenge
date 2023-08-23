# VBA-challenge

Sub AnalyzeStockData():

    For Each Worksheet In Worksheets
    
        Dim sheetName As String
        Dim currentRow As Long
        Dim startOfTickerBlock As Long
        Dim tickerCounter As Long
        Dim lastRowColumnA As Long
        Dim lastRowColumnI As Long
        Dim percentChange As Double
        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        Dim greatestVolume As Double
        
        ' Get the name of the worksheet
        sheetName = Worksheet.Name
        
        ' Create column headers
        Worksheet.Cells(1, 9).Value = "Ticker"
        Worksheet.Cells(1, 10).Value = "Yearly Change"
        Worksheet.Cells(1, 11).Value = "Percent Change"
        Worksheet.Cells(1, 12).Value = "Total Stock Volume"
        Worksheet.Cells(1, 16).Value = "Ticker"
        Worksheet.Cells(1, 17).Value = "Value"
        Worksheet.Cells(2, 15).Value = "Greatest % Increase"
        Worksheet.Cells(3, 15).Value = "Greatest % Decrease"
        Worksheet.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Initialize ticker counter and start row
        tickerCounter = 2
        startOfTickerBlock = 2
        
        ' Find the last non-blank cell in column A
        lastRowColumnA = Worksheet.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Loop through all rows
        For currentRow = 2 To lastRowColumnA
        
            ' Check if ticker name changed
            If Worksheet.Cells(currentRow + 1, 1).Value <> Worksheet.Cells(currentRow, 1).Value Then
            
                ' Write ticker in column I (#9)
                Worksheet.Cells(tickerCounter, 9).Value = Worksheet.Cells(currentRow, 1).Value
                
                ' Calculate and write Yearly Change in column J (#10)
                Worksheet.Cells(tickerCounter, 10).Value = Worksheet.Cells(currentRow, 6).Value - Worksheet.Cells(startOfTickerBlock, 3).Value
                
                ' Conditional formatting
                If Worksheet.Cells(tickerCounter, 10).Value < 0 Then
                    ' Set cell background color to red
                    Worksheet.Cells(tickerCounter, 10).Interior.ColorIndex = 3
                Else
                    ' Set cell background color to green
                    Worksheet.Cells(tickerCounter, 10).Interior.ColorIndex = 4
                End If
                
                ' Calculate and write percent change in column K (#11)
                If Worksheet.Cells(startOfTickerBlock, 3).Value <> 0 Then
                    percentChange = ((Worksheet.Cells(currentRow, 6).Value - Worksheet.Cells(startOfTickerBlock, 3).Value) / Worksheet.Cells(startOfTickerBlock, 3).Value)
                    Worksheet.Cells(tickerCounter, 11).Value = Format(percentChange, "Percent")
                Else
                    Worksheet.Cells(tickerCounter, 11).Value = Format(0, "Percent")
                End If
                
                ' Calculate and write total volume in column L (#12)
                Worksheet.Cells(tickerCounter, 12).Value = WorksheetFunction.Sum(Range(Worksheet.Cells(startOfTickerBlock, 7), Worksheet.Cells(currentRow, 7)))
                
                ' Increment ticker counter
                tickerCounter = tickerCounter + 1
                
                ' Set new start row of the ticker block
                startOfTickerBlock = currentRow + 1
                
            End If
        
        Next currentRow
        
        ' Find last non-blank cell in column I
        lastRowColumnI = Worksheet.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Initialize variables for summary
        greatestVolume = Worksheet.Cells(2, 12).Value
        greatestIncrease = Worksheet.Cells(2, 11).Value
        greatestDecrease = Worksheet.Cells(2, 11).Value
        
        ' Loop for summary
        For currentRow = 2 To lastRowColumnI
        
            ' Check for greatest total volume
            If Worksheet.Cells(currentRow, 12).Value > greatestVolume Then
                greatestVolume = Worksheet.Cells(currentRow, 12).Value
                Worksheet.Cells(4, 16).Value = Worksheet.Cells(currentRow, 9).Value
            End If
            
            ' Check for greatest increase
            If Worksheet.Cells(currentRow, 11).Value > greatestIncrease Then
                greatestIncrease = Worksheet.Cells(currentRow, 11).Value
                Worksheet.Cells(2, 16).Value = Worksheet.Cells(currentRow, 9).Value
            End If
            
            ' Check for greatest decrease
            If Worksheet.Cells(currentRow, 11).Value < greatestDecrease Then
                greatestDecrease = Worksheet.Cells(currentRow, 11).Value
                Worksheet.Cells(3, 16).Value = Worksheet.Cells(currentRow, 9).Value
            End If
            
            ' Write summary results
            Worksheet.Cells(2, 17).Value = Format(greatestIncrease, "Percent")
            Worksheet.Cells(3, 17).Value = Format(greatestDecrease, "Percent")
            Worksheet.Cells(4, 17).Value = Format(greatestVolume, "Scientific")
        
        Next currentRow
        
        ' Adjust column width automatically
        Worksheets(sheetName).Columns("A:Z").AutoFit
            
    Next Worksheet
        
End Sub
