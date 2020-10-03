Attribute VB_Name = "Module1"
Sub stock_analysis():
  For Each WS In Worksheets
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
    ' label for new columns
        WS.Range("I1").Value = "Ticker"
        
        WS.Range("J1").Value = "Yearly change"
        
       WS.Range("K1").Value = "Percent change"
        
       WS.Range("L1").Value = "Total stock volume"
       
       WS.Range("O2").Value = "Greatest % increase"
       
       WS.Range("O3").Value = "Greatest % decrease"
       
       WS.Range("O4").Value = "Greatest total volume"
       
       WS.Range("P1").Value = "Ticker"
       
       WS.Range("Q1").Value = "Value"
        
        WS.Range("Q2:Q3").NumberFormat = "0.00%"
     
 ' set last row
        last_row = WS.Cells(Rows.Count, 1).End(xlUp).Row
        ' set i for master row iterator
        For i = 2 To last_row
        'set j for summary table iterator
             For j = 2 To last_row
 'out putting ticker symbol and total stock volume
     If WS.Cells(j, 9).Value = WS.Cells(i, 1).Value Then
        WS.Cells(j, 12).Value = WS.Cells(j, 12).Value + WS.Cells(i, 7).Value
             
        Exit For
                    
    ElseIf WS.Cells(j, 9).Value = "" Then
        WS.Cells(j, 9).Value = WS.Cells(i, 1).Value
        WS.Cells(j, 12).Value = WS.Cells(i, 7).Value
                    
        Exit For
    End If
    Next j
    'set yearly change
    If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
             open_price = WS.Cells(ticker_row, 3).Value
             closing_price = WS.Cells(i, 6).Value
             yearly_change = closing_price - open_price
             WS.Cells(Summary_table, 10).Value = yearly_change
            
            'calculate & set percent change and if statement for if open price is less than 0
             If open_price <= 0 Then
                percent_change = 0
            Else
            percent_change = yearly_change / open_price
            WS.Cells(Summary_table, 11).Value = percent_change
            WS.Cells(Summary_table, 11).NumberFormat = "0.00%"
            End If
         
         ' highlight positives and negative values in yearly change
         If WS.Cells(Summary_table, 10).Value >= 0 Then
            WS.Cells(Summary_table, 10).Interior.Color = vbGreen
         Else
            WS.Cells(Summary_table, 10).Interior.Color = vbRed
         End If
         'add 1 to summary table
         Summary_table = Summary_table + 1
         ticker_row = i + 1
    End If
       Next i
 
 'challenge part
last_row = WS.Cells(Rows.Count, 9).End(xlUp).Row
For i = 2 To last_row

' set variables
ticker = WS.Cells(i, 9).Value
percent_change = WS.Cells(i, 11).Value
total_volume = WS.Cells(i, 12).Value
          
  'match percent change to greatest % increase
   If percent_change > WS.Range("Q2").Value Then
        WS.Range("Q2").Value = percent_change
        WS.Range("P2").Value = ticker
   End If
 'match percent change to greatest % decrease
  If percent_change < WS.Range("Q3").Value Then
        WS.Range("Q3").Value = percent_change
        WS.Range("P3").Value = ticker
  End If
 ' match total volume to greatest volume
  If total_volume > WS.Range("Q4").Value Then
        WS.Range("Q4").Value = total_volume
        WS.Range("P4").Value = ticker
  End If
    Next i
         Next WS

End Sub
