Attribute VB_Name = "Module1"
Sub Stock_info()

For Each ws In Worksheets
     'worksheetname
     WorkshetName = ws.Name
     
     'Keep track of each stock summary
     Summary_Row = 2
     
     'Set opening price
     open_price = ws.Cells(2, 3).Value
     
     'Set a volume
     Volume = 0
     
     'Summary table headers
     ws.Cells(1, 9).Value = "Ticker"
     ws.Cells(1, 10).Value = "Yearly Change"
     ws.Cells(1, 11).Value = "Percent Change"
     ws.Cells(1, 12).Value = "Total Stock Volume"
     
     'Last row of stock summary table
     stock_sum_lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
     
     'Last row of ticker column
     lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
     
     'Loop through ticker
     For i = 2 To lastrow
         
         'Check if same ticker
         If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
         
            'Set ticker variable
             Ticker = ws.Cells(i, 1).Value
         
             'Add to Volume
             Volume = Volume + ws.Cells(i, "G").Value
             
             'Print ticker name in Summary row
             ws.Cells(Summary_Row, 9).Value = Ticker
             
             'Print Total Stock
             ws.Cells(Summary_Row, 12).Value = Volume
             ws.Cells(Summary_Row, 12).NumberFormat = "General"
             
             'Grab value of closing  column, store it, & subtract it from opening
             Closing_price = ws.Cells(i, "F").Value
             yearly_change = (Closing_price - open_price)
             
             'Print yearly change
             ws.Cells(Summary_Row, 10).Value = yearly_change
             
             'Calculate percent change
             percent_change = (yearly_change / open_price)
                  
             
             'Display percent change
             ws.Cells(Summary_Row, 11).Value = percent_change
             ws.Cells(Summary_Row, 11).NumberFormat = "0.00%"
             
             
             'Add one to ticker output
             Summary_Row = Summary_Row + 1
             
             'Reset Volume Total
             Volume = 0
         
             'Reset opening price
             open_price = ws.Cells(i + 1, 3).Value
             
         Else
         
             'Add to volume total
             Volume = Volume + ws.Cells(i, "G").Value
             
         End If
    Next i
    
    'percentage cell fill
      For i = 2 To stock_sum_lastrow
         If ws.Cells(i, "K").Value < 0 Then
             ws.Cells(i, "K").Interior.ColorIndex = 3
         Else
             ws.Cells(i, "K").Interior.ColorIndex = 4
         End If
     Next i
     
Next ws
End Sub
Sub percent_()
For Each ws In Worksheets
     'worksheetname
     WorkshetName = ws.Name
    
     'Summary Percentage table headers
     ws.Range("P1").Value = "Ticker"
     ws.Range("Q1").Value = "Value"
     ws.Range("O2").Value = "Greatest % Increase"
     ws.Range("O3").Value = "Greatest % Decrease"
     ws.Range("O4").Value = "Greatest Total Volume"
     
     'Last row of stock summary table
     stock_sum_lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
     
     'Loop stock summary table
     For i = 2 To stock_sum_lastrow
        'Find max
        If ws.Cells(i, "K").Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & stock_sum_lastrow)) Then
            ws.Cells(2, "P").Value = ws.Cells(i, 9).Value
            ws.Cells(2, "Q").Value = ws.Cells(i, 11).Value
            ws.Cells(2, "Q").NumberFormat = "0.00%"
        
        'Find min
        ElseIf ws.Cells(i, "K").Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & stock_sum_lastrow)) Then
            ws.Cells(3, "P").Value = ws.Cells(i, 9).Value
            ws.Cells(3, "Q").Value = ws.Cells(i, 11).Value
            ws.Cells(3, "Q").NumberFormat = "0.00%"
           
        'Max total vol
        ElseIf ws.Cells(i, "L").Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & stock_sum_lastrow)) Then
            ws.Cells(4, "P").Value = ws.Cells(i, 9).Value
            ws.Cells(4, "Q").Value = ws.Cells(i, 12).Value
        End If
    
    Next i
    
Next ws
End Sub
