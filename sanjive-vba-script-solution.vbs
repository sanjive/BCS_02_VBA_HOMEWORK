Sub ws_TotalVolume()
    'The routine calculates the total volume of shares traded for the
    'given ticker for the year
	' Sanjive Agarwal 
	' March 13, 2020
    
    Dim ws_TotalVolume As LongLong
    Dim ws_FirstRow, ws_LastRow, tkr_FirstRow As Long
    Dim ws_count As Integer
    Dim ws As Worksheet
    Dim wb As Workbook: Set wb = ThisWorkbook
        
    Dim Ticker As String
    Dim Volume, PriceOpen, PriceClose, PriceLow, PriceHigh As Double
    
    Dim tkr_Days, tkr_ResultRowCtr As Integer
    Dim tkr_Change, tkr_DailyChange, tkr_PctChange As Double
    
    'Get the worksheet names in the array
    For Each ws In wb.Worksheets
        'Select the current worksheet
        ws.Select
        
        'Initialize the variables for the selected worksheet
        ws_TotalVolume = 0          'Total for the Stock Volume for a ticker in one worksheet
        ws_FirstRow = 2             'Start of the row position in a worksheet
        ws_LastRow = Cells(Rows.Count, 1).End(xlUp).Row    'End position of the worksheet
        tkr_ResultRowCtr = 2        'Counter to track the result rows
        
        tkr_FirstRow = 2            'Start of the Ticker row position in the worksheet
        
        'Write the column header for the calculated data
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Average Change"
        Cells(1, 12).Value = "% Change"
        Cells(1, 13).Value = "Total Stock Volume"
        Cells(1, 14).Value = "Price Trend"
        
        Cells(1, 9).Font.Bold = True
        Cells(1, 10).Font.Bold = True
        Cells(1, 11).Font.Bold = True
        Cells(1, 12).Font.Bold = True
        Cells(1, 13).Font.Bold = True
        Cells(1, 14).Font.Bold = True

        'Autofit the columns width
        Range("I:N").EntireColumn.AutoFit
        

        'Process each Ticker
        For I = ws_FirstRow To ws_LastRow
        
            'Get the ticker for which the data is being processed
            Ticker = Cells(I, 1).Value
            
            If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
                'Processing the last record of the ticker in the list of sorted ticker data
            
                'Populate the result data
                PriceOpen = Cells(tkr_FirstRow, 3)  'Open Price at the begining of the year
                PriceClose = Cells(I, 6)            'Close Price at the end of year
                PriceLow = Cells(I, 5).Value
                PriceHigh = Cells(I, 4).Value
                Volume = Cells(I, 7).Value
                
                tkr_Change = PriceClose - PriceOpen
                
                'Skip the division if the divisor is Zero
                If PriceOpen <> 0 Then
                    'Calculate as the ratio and the days; cell data format will convert to %age
                    'This is the total change from open price to close price
                    tkr_PctChange = tkr_Change / PriceOpen
                End If
                'Below the stock's daily low and high price are agreegated for the year
                tkr_DailyChange = tkr_DailyChange + (PriceHigh - PriceLow)
                ws_TotalVolume = ws_TotalVolume + Volume
                
                'Number of days of ticker data to be used for calculating Average
                'Skip the division if the divisor is Zero
                tkr_Days = (I - tkr_FirstRow) + 1
                If tkr_Days <> 0 Then
                    tkr_AvgChange = tkr_Change / tkr_Days
                End If
                
                'Write the values to the target location
                Cells(tkr_ResultRowCtr, 9).Value = Ticker
                Cells(tkr_ResultRowCtr, 10).Value = tkr_Change
                'Cells(tkr_ResultRowCtr, 10).NumberFormat = "[Color 10]#,##0.00;[Red]-#,##0.00"
                Cells(tkr_ResultRowCtr, 10).NumberFormat = "#,##0.00"
                Cells(tkr_ResultRowCtr, 11).Value = Format(tkr_AvgChange, "#,##0.00000000")
                Cells(tkr_ResultRowCtr, 12).Value = Format(tkr_PctChange, "##0.00%")
                Cells(tkr_ResultRowCtr, 13).Value = Format(ws_TotalVolume, "#,##0")
                
                'The below code writes a green or Red Triangle depending on the price direction for the
                ' ticker at the end of the year
                If tkr_Change > 0 Then
                    Cells(tkr_ResultRowCtr, 14).Value = ChrW(11205)  'Up Triangle Symbol
                    Cells(tkr_ResultRowCtr, 14).NumberFormat = "[Color 10]" 'Green
                ElseIf tkr_Change < 0 Then
                    Cells(tkr_ResultRowCtr, 14).Value = ChrW(11206) 'Down Triangle Symbol
                    Cells(tkr_ResultRowCtr, 14).NumberFormat = "[Red]"
                Else
                    Cells(tkr_ResultRowCtr, 14).Value = ChrW(11208) 'Side Triangle symbol
                    Cells(tkr_ResultRowCtr, 14).NumberFormat = "[Black]"
                End If
                
                'Apply conditional formatting to results data
                Select Case tkr_Change
                    Case Is > 0
                        ws.Cells(tkr_ResultRowCtr, 10).Interior.ColorIndex = 4
                    Case Is < 0
                        ws.Cells(tkr_ResultRowCtr, 10).Interior.ColorIndex = 3
                    Case Else
                        ws.Cells(tkr_ResultRowCtr, 10).Interior.ColorIndex = 0
                End Select
                
                'Reset the Variables as the Ticker changes
                ws_TotalVolume = 0
                tkr_Change = 0
                tkr_ResultRowCtr = tkr_ResultRowCtr + 1
                tkr_PctChange = 0
                tkr_Days = 0
                tkr_DailyChange = 0
                
                'Set the Start row for the next ticker by incrementing the row counter "i"
                tkr_FirstRow = I + 1
                
            Else
                'Price of the Ticker on the first day
                PriceOpen = Cells(tkr_FirstRow, 3)  'Price Open value at the begin of the year
                'Price of the ticker on curent day as indicated by the variable i
                PriceClose = Cells(I, 6)            'Price Close value on the current day
                PriceLow = Cells(I, 5).Value
                PriceHigh = Cells(I, 4).Value
                Volume = Cells(I, 7).Value
                
                'The day's stock volume is added here to the total
                ws_TotalVolume = ws_TotalVolume + Volume
                'Below the stock's daily low and high price are aggregated for the year
                tkr_DailyChange = tkr_DailyChange + (PriceHigh - PriceLow)

            End If
            
        'Next Ticker
        Next I
                
        'Take the Max and min of the result data (Columns I thru N in the spreadsheet)
        'Set the Header values
        ws.Cells(1, 17) = "Ticker"
        ws.Cells(1, 18) = "Value"
        ws.Cells(2, 16) = "Greatest % Increase:"
        ws.Cells(3, 16) = "Greatest % Decrease:"
        ws.Cells(4, 16) = "Greatest Total Volume:"
        'Set the font Bold for the headers
        ws.Cells(1, 17).Font.Bold = True
        ws.Cells(1, 18).Font.Bold = True
        ws.Cells(1, 18).Font.Bold = True
        ws.Cells(2, 16).Font.Bold = True
        ws.Cells(3, 16).Font.Bold = True
        ws.Cells(4, 16).Font.Bold = True
        
        resultLastRow = tkr_ResultRowCtr - 1
        
        'Format the data type of the cells being written
        'Maximum % increase
        ws.Cells(2, 18) = Format(WorksheetFunction.Max(ws.Range("L2:L" & resultLastRow)), "#,000.00%")
        'Minimum % Decrease
        ws.Cells(3, 18) = Format(WorksheetFunction.Min(ws.Range("L2:L" & resultLastRow)), "#,000.00%")
        'Maximum Volume
        ws.Cells(4, 18) = Format(WorksheetFunction.Max(ws.Range("M2:M" & resultLastRow)), "#,##0")
        
        'Find the cell with the value identified in the above written cell so that the correspoding Ticker can be identified
        MaxIncrementRow = WorksheetFunction.Match(ws.Cells(2, 18), ws.Range("L2:L" & resultLastRow), 0)
        MaxDecrementRow = WorksheetFunction.Match(ws.Cells(3, 18), ws.Range("L2:L" & resultLastRow), 0)
        MaxVolumeRow = WorksheetFunction.Match(ws.Cells(4, 18), ws.Range("M2:M" & resultLastRow), 0)
        
        ws.Cells(2, 17) = ws.Cells(MaxIncrementRow + 1, 9)
        ws.Cells(3, 17) = ws.Cells(MaxDecrementRow + 1, 9)
        ws.Cells(4, 17) = ws.Cells(MaxVolumeRow + 1, 9)
       
        ws.Range("P:R").EntireColumn.AutoFit
                           
    'Next Worksheet
    Next ws
    
End Sub
