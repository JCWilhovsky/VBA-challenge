Attribute VB_Name = "Module1"
Sub StockAnalysis()
'Set variables
    Dim Total As Double
    Dim row As Long
    Dim rowCount As Double
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim summaryTableRow As Long
    Dim stockStartRow As Long
    Dim startValue As Long
    Dim lastTicker As String
    
'Loop through all the sheets in the excel workbook
    For Each ws In Worksheets
    

'Set Title Row for the new columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
'Values for the summary table row
    summaryTableRow = 0
    Total = 0
    quartleryChange = 0
    stockStartRow = 2
    startValue = 2

    rowCount = ws.Cells(Rows.Count, "A").End(xlUp).row
    lastTicker = ws.Cells(rowCount, 1).Value
    For row = 2 To rowCount
    
'check for any changes in the tickers
    If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
    
'If there is a change in the Column A
'add to the total stock volume one last time
    Total = Total + ws.Cells(row, 7).Value

'check to see if the value of the total stock volume is 0
    If Total = 0 Then
    ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value  'Prints ticker value from colum A
    ws.Range("J" & 2 + summaryTableRow).Value = 0   'prints a 0 in column J (Quarterly chnage)
    ws.Range("K" & 2 + summaryTableRow).Value = 0   'prints a 0 in column K (Percent change)
    ws.Range("L" & 2 + summaryTableRow).Value = 0   'prints a 0 in column L (Total stock volume)

    
    Else
        'find the first non-zero first open value for the stock
    If ws.Cells(startValue, 3).Value = 0 Then
        'if the first open is 0, search for the first non-zero stock open value by moving to the next rows
        For findValue = startValue To row
        
        If ws.Cells(findValue, 3).Value <> 0 Then
            startValue = findValue
            'finally break from loop
            Exit For
         End If
        
        Next findValue
    End If
        
        
        'calculate the quarterly change (difference in the last close and first open)
        quarterlyChange = ws.Cells(row, 6).Value - ws.Cells(startValue, 3).Value
    
        'calculate the percent change (quarterly change / first open)
        percentChange = quarterlyChange / ws.Cells(startValue, 3).Value
        'Print the results
        ws.Range("I" & 2 + summaryTableRow).Value = Cells(row, 1).Value
        ws.Range("J" & 2 + summaryTableRow).Value = quarterlyChange
        ws.Range("K" & 2 + summaryTableRow).Value = percentChange
        ws.Range("L" & 2 + summaryTableRow).Value = Total
        
        
        'Color the quarterly change column in the summary based on our values
        If quarterlyChange > 0 Then
            ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 4
        ElseIf quarterlyChange < 0 Then
            ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 3
        Else
            ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 0
        End If
        
        'reset / update the values for the next ticker
        Total = 0
        averageChange = 0
        quarterlyChange = 0
        startValue = row + 1
        summaryTableRow = summaryTableRow + 1
        
      End If
        
    
    Else

    Total = Total + ws.Cells(row, 7).Value

      End If
       
    
    Next row
  
    summaryTableRow = ws.Cells(Rows.Count, "I").End(xlUp).row
      
    'find the last data in the extra rows from columns J-L
    Dim lastExtraRow As Long
    lastExtra = ws.Cells(Rows.Count, "J").End(xlUp).row
      
    'loop that clears the extra data from columns I-L
        For e = summaryTableRow To lastExtraRow
          For Column = 9 To 12
            ws.Cells(e, Column).Value = ""
            ws.Cells(e, Column).Interior.ColorIndex = 0
            
        Next Column
    Next e
            
    ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & summaryTableRow + 2))
    ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & summaryTableRow + 2))
    ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & summaryTableRow + 2))
    
    Dim greatestIncreaseRow As Double
    Dim greatestDecreaseRow As Double
    Dim greatestTotVolRow As Double
    greatestIncreaseRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & summaryTableRow + 2)), ws.Range("K2:K" & summaryTableRow + 2), 0)
    greatestDecreaseRow = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & summaryTableRow + 2)), ws.Range("K2:K" & summaryTableRow + 2), 0)
    greatestTotVolRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & summaryTableRow + 2)), ws.Range("L2:L" & summaryTableRow + 2), 0)
    
    
    'show ticker symbol for the greatest increase, greatest decrease, greatest total stock volume
    ws.Range("P2").Value = ws.Cells(greatestIncreaseRow + 1, 9).Value
    ws.Range("P3").Value = ws.Cells(greatestDecreaseRow + 1, 9).Value
    ws.Range("P4").Value = ws.Cells(greatestTotVolRow + 1, 9).Value
    
    'format the summary table columns
    For s = 0 To summaryTableRow
        ws.Range("J" & 2 + s).NumberFormat = "0.00"
        ws.Range("K" & 2 + s).NumberFormat = "0.00%"
        ws.Range("L" & 2 + s).NumberFormat = "#.###"
        
    Next s
    
    'Format the summary aggregate
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("Q4").NumberFormat = "#,###"
    
    
'Fix/condition how the title in the columns look
    ws.Columns("A:Q").AutoFit
    
    
    Next ws
    
    
    
End Sub
