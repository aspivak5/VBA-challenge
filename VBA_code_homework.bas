
Sub stock_analysis():
  For Each ws In Worksheets
    Dim ticker As String
    Dim last_row As Long
    Dim Summary_table As Long
        Summary_table = 2
    Dim open_price As Double
    Dim closing_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim ticker_row As Long
     ticker_row = 2
    Dim max_percent As Double
    Dim min_percent As Double
    Dim max_volume As Double
    ' label for new columns
        ws.Range("I1").Value = "Ticker"
        
        ws.Range("J1").Value = "Yearly change"
        
       ws.Range("K1").Value = "Percent change"
        
       ws.Range("L1").Value = "Total stock volume"
       
       ws.Range("O2").Value = "Greatest % increase"
       
       ws.Range("O3").Value = "Greatest % decrease"
       
       ws.Range("O4").Value = "Greatest total volume"
       
       ws.Range("P1").Value = "Ticker"
       
       ws.Range("Q1").Value = "Value"
        
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
     
 ' set last row
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ' set i for master row iterator
        For i = 2 To last_row
        'set j for summary table iterator
             For j = 2 To last_row
 'out putting ticker symbol and total stock volume
     If ws.Cells(j, 9).Value = ws.Cells(i, 1).Value Then
        ws.Cells(j, 12).Value = ws.Cells(j, 12).Value + ws.Cells(i, 7).Value
             
        Exit For
                    
    ElseIf ws.Cells(j, 9).Value = "" Then
        ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
        ws.Cells(j, 12).Value = ws.Cells(i, 7).Value
                    
        Exit For
    End If
    Next j
    'set yearly change
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
             open_price = ws.Cells(ticker_row, 3).Value
             closing_price = ws.Cells(i, 6).Value
             yearly_change = closing_price - open_price
             ws.Cells(Summary_table, 10).Value = yearly_change
            
            'calculate & set percent change and if statement for if open price is less than 0
             If open_price <= 0 Then
                percent_change = 0
            Else
            percent_change = yearly_change / open_price
            ws.Cells(Summary_table, 11).Value = percent_change
            ws.Cells(Summary_table, 11).NumberFormat = "0.00%"
            End If
         
         ' highlight positives and negative values in yearly change
         If ws.Cells(Summary_table, 10).Value >= 0 Then
            ws.Cells(Summary_table, 10).Interior.Color = vbGreen
         Else
            ws.Cells(Summary_table, 10).Interior.Color = vbRed
         End If
         'add 1 to summary table
         Summary_table = Summary_table + 1
         ticker_row = i + 1
    End If
       Next i
 
 'challenge part
For i = 2 To last_row
last_row = ws.Cells(Rows.Count, 9).End(xlUp).Row
' set variables
ticker = ws.Cells(i, 9).Value
percent_change = ws.Cells(i, 11).Value
total_volume = ws.Cells(i, 12).Value
'set max/min % increase and greatest volume increase
max_percent = Application.WorksheetFunction.Max(ws.Range("K2:K" & last_row))
min_percent = Application.WorksheetFunction.Min(ws.Range("K2:K" & last_row))
max_volume = Application.WorksheetFunction.Max(ws.Range("L2:L" & last_row))
          
  'match percent change to greatest % increase
   If percent_change = max_percent Then
     ws.Range("P2").Value = ticker
     ws.Range("Q2").Value = percent_change
   End If
 'match percent change to greatest % decrease
  If percent_change = min_percent Then
        ws.Range("Q3").Value = percent_change
        ws.Range("P3").Value = ticker
  End If
 ' match total volume to greatest volume
  If total_volume = max_volume Then
        ws.Range("Q4").Value = total_volume
        ws.Range("P4").Value = ticker
  End If
    Next i
         Next ws

End Sub
