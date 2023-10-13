Attribute VB_Name = "Module1"
Sub stock_market()


'Declare and set variables'
Dim ws As Worksheet
Dim r As Long
Dim summary_ticker_row As Integer
Dim stock_volume As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim newticker As Long
Dim Lastrow As Long
Dim days As Integer
Dim dailychange As Double
Dim averagechange As Double


'create a script that loops through all stocks'
For Each ws In Worksheets
    yearly_change = 0
    stock_volume = 0
    newticker = 2
    summary_ticker_row = 0

'create columns'
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Value"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
    

'Define lastrow of worksheet'
    Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Do loop of current worksheet to Lastrow
    For r = 2 To Lastrow

    'Ticker symbol changed
    If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then
        
        'stock volume
        stock_volume = stock_volume + Cells(r, 7).Value
        
        If stock_volume = 0 Then
            ws.Range("I" & 2 + summary_ticker_row).Value = Cells(r, 1).Values
            ws.Range("J" & 2 + summary_ticker_row).Value = 0
            ws.Range("K" & 2 + summary_ticker_row).Value = "%" & 0
            ws.Range("L" & 2 + summary_ticker_row).Value = 0
            
        'worked with tutor***
        Else
            If Cells(newticker, 3).Value = 0 Then
                For find_value = newticker To r
                    If Cells(find_value, 3).Value <> 0 Then
                        newticker = find_value
                        Exit For
                    End If
                Next find_value
            End If
    
        
        'Yearly change
            yearly_change = (Cells(r, 6).Value - Cells(newticker, 3).Value)
            percent_change = (yearly_change) / (Cells(newticker, 3).Value)

        'restart ticker
        newticker = r + 1
        
        'print in the summary table
            ws.Range("I" & 2 + summary_ticker_row).Value = Cells(r, 1).Value
            ws.Range("J" & 2 + summary_ticker_row).Value = yearly_change
            ws.Range("J" & 2 + summary_ticker_row).NumberFormat = "0.00"
            ws.Range("K" & 2 + summary_ticker_row).Value = percent_change
            ws.Range("K" & 2 + summary_ticker_row).NumberFormat = "0.00%"
            ws.Range("L" & 2 + summary_ticker_row).Value = stock_volume
            
        'conditional formatting that will highlight positive and negative changes'

            lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
       
    
        
End If
        
        'reset variables
        summary_ticker_row = summary_ticker_row + 1
        stock_volume = 0
        yearly_change = 0
        days = 0
    
        
        Else
            'add volume
            stock_volume = stock_volume + ws.Cells(r, 7).Value
        
    End If
Next r

For r = 2 To lastrow_summary_table
    If Cells(r, 10).Value > 0 Then
        Cells(r, 10).Interior.ColorIndex = 4
    Else
        Cells(r, 10).Interior.ColorIndex = 3
End If
Next r

'Worked with tutor on this section***
'Greatest % increase,decrease, and greatest total volume
ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & Lastrow)) * 100
ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & Lastrow)) * 100
ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & Lastrow))

greatest_increase = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & Lastrow)), ws.Range("K2:K" & Lastrow), 0)
greatest_decrease = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & Lastrow)), ws.Range("K2:K" & Lastrow), 0)
greatest_volume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & Lastrow)), ws.Range("L2:L" & Lastrow), 0)

'Ticker symbol for greatest % increase, decrease, and greatest total volume
ws.Range("P2") = Cells(greatest_increase + 1, 9)
ws.Range("P3") = Cells(greatest_increase + 1, 9)
ws.Range("P4") = Cells(greatest_volume + 1, 9)

Next ws
End Sub




