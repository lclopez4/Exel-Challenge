Sub VBS_Challenge

' Initiate variables to roll through the Sheets
Dim wb As Workbook
Dim ws As Worksheet

'get active workbook
Set wb = ActiveWorkbook

'loop through all the sheets, one at a time, and build the summary table and "podium" on each sheet
'the way this is done is fun to watch as the summary cells get upddated for each line in the 'data' table
'BUT IT IS PRETTY RESOURCE INTENSIVE! YOU MAY NOT WANT THIS IN A PRODUCTION ENVIRONMENT.
For w = 1 To wb.Worksheets.Count
'This is the start of the 'single sheet' workflow for each sheet (tab) it finds
    
    'activate the current sheet
    wb.Worksheets(w).Activate
    
        
        Dim ticker As String
        Dim tickerCount As Integer
        Dim startPrice As Double
        Dim endPrice As Double
        Dim priceChangePercent As Double
        Dim volume As Double
        Dim maxIncrease As Double
        Dim maxIncreaseTicker As String
        Dim minIncrease As Double
        Dim minIncreaseTicker As String
        Dim maxVolume As Double
        Dim maxVolumeTicker As String
        Dim tabName As String
        
        
        'set the stage, so to speak
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Change in price"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase:"
        Cells(3, 15).Value = "Greatest % Decrease:"
        Cells(4, 15).Value = "Greatest Total Volume:"
        
        Columns("A:G").ColumnWidth = 10
        Columns("H").ColumnWidth = 2
        Columns("I").ColumnWidth = 10
        Columns("J:K").ColumnWidth = 13
        Columns("L").ColumnWidth = 15
        Columns("M:N").ColumnWidth = 2
        Columns("O").ColumnWidth = 20
        Columns("P").ColumnWidth = 10
        Columns("Q").ColumnWidth = 15
        
        Columns("A:Q").HorizontalAlignment = xlCenter
        Columns("O").HorizontalAlignment = xlRight
        
        'Start the Process of building the Summary Table
        tickerCount = 1
        
        'loop through the rows and pull the rows
        For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
            'MsgBox ("there are " & Str(Cells(Rows.Count, 1).End(xlUp).Row) + " Rows")
        
            'MsgBox ("number of columns: " + Str(Cells(1, Columns.Count).End(xlToLeft).Column))
        
                'check if its a new ticker:
                'The "-1" below ensures the first new ticker is found, not the last
                If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                    'MsgBox ("new Ticker!")
                    'advance the Summary Rows (counts unique ticker symbols)
                    tickerCount = tickerCount + 1
                    
                    'register the ticker
                    ticker = Cells(i, 1).Value
                    
                    'register the open price
                    startPrice = Cells(i, 3).Value
                    'MsgBox (Str(startPrice))
                    
                    'put the ticker in column 9
                    Cells(tickerCount, 9).Value = ticker
         
                    'register the starting volume
                    volume = Cells(i, 7).Value
                    'set the starting volume in column 12
                    Cells(tickerCount, 12).Value = volume
                
                'if the ticker is the same:
                Else
                    'register the end price
                    endPrice = Cells(i, 6).Value
                    
                    'keep adding the volume as long as the ticker hasn't changed
                    'MsgBox ("This is the row's volume: " + Str(Cells(i, 7).Value))
                    volume = volume + Cells(i, 7).Value
                    'reset the volume value in Column 12 (fun to watch)
                    Cells(tickerCount, 12).Value = volume
                    
        
                    'put the calculated change in column 10
                    Cells(tickerCount, 10).Value = endPrice - startPrice
                    
                    'change the color of the cell, based on the change (also fun to watch)
                    If Cells(tickerCount, 10).Value < 0 Then
                        Cells(tickerCount, 10).Interior.ColorIndex = 3
                    Else
                        Cells(tickerCount, 10).Interior.ColorIndex = 4
                    End If
                    
                    'calculate the percent change
                    priceChangePercent = (endPrice - startPrice) / startPrice
                    'put the calculated pecent change in Column 11
                    Cells(tickerCount, 11).Value = priceChangePercent
                    
                    'format the cells's properly
                    Cells(tickerCount, 10).NumberFormat = "$#,##0.00"
                    Cells(tickerCount, 11).NumberFormat = "0.00%"
                    Cells(tickerCount, 12).NumberFormat = "#,###"
                    
                
                End If
        
        Next i
        
        'initialize variables to get the winners and loser
        maxIncrease = 0
        minIncrease = 0
        maxVolume = 0
        'now loop through the summery and find the big winner and the big loser
        'Start "podium work"
        For x = 2 To Cells(Rows.Count, 9).End(xlUp).Row
            
                'look for the greatest increase percentage
                If Cells(x, 11).Value > maxIncrease Then
                    maxIncrease = Cells(x, 11).Value
                    maxIncreaseTicker = Cells(x, 9).Value
                End If
                
                'look for the greatest decrease percentage
                If Cells(x, 11).Value < minIncrease Then
                    minIncrease = Cells(x, 11).Value
                    minIncreaseTicker = Cells(x, 9).Value
                End If
            
                'look for the max total volume
                If Cells(x, 12).Value > maxVolume Then
                    maxVolume = Cells(x, 12).Value
                    maxVolumeTicker = Cells(x, 9).Value
                End If
        'end of "podium" work
        Next x
        
        'set the ticker and values
            Cells(2, 16).Value = maxIncreaseTicker
            Cells(2, 17).Value = maxIncrease
            Cells(2, 17).NumberFormat = "0.00%"
            Cells(3, 16).Value = minIncreaseTicker
            Cells(3, 17).Value = minIncrease
            Cells(3, 17).NumberFormat = "0.00%"
            Cells(4, 16).Value = maxVolumeTicker
            Cells(4, 17).Value = maxVolume
            Cells(4, 17).NumberFormat = "#,###"
        
        
        
        'END SINGLE SHEET WORK


Next w



End Sub
