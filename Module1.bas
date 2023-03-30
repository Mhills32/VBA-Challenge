Attribute VB_Name = "Module1"

Sub stock_ticker()

    Dim ws As Worksheet
    Dim ticker As String
    Dim rowcount As Long
    Dim total_volume As Double
    total_volume = 0
    Dim i As Double
    Dim row_tracker As Double
    Dim start As Long
    Dim maxValue As Double
    Dim minValue As Double
    Dim ticker1 As String
    Dim ticker3 As String
    
    For Each ws In ThisWorkbook.Worksheets
    
        row_tracker = 2
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Year Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        rowcount = ws.Cells(Rows.Count, 1).End(xlUp).Row
        start = 2
        maxValue = 0
        minValue = 0
        maxVolume = 0
  
        For i = 2 To rowcount
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                total_volume = total_volume + ws.Cells(i, 7).Value
                ws.Range("I" & row_tracker).Value = ticker
                ws.Range("L" & row_tracker).Value = total_volume
                
                If Cells(start, 3) = 0 Then
                    For first_nonzero = start To i
                        If Cells(first_nonzero, 3).Value <> 0 Then
                            start = first_nonzero
                            Exit For
                        End If
                    Next first_nonzero
                End If
                
                year_change = ws.Cells(i, 6) - ws.Cells(start, 3)
                percent_change = Round(year_change / ws.Cells(start, 3) * 100, 2)
                ws.Range("J" & row_tracker).Value = year_change
                ws.Range("K" & row_tracker).Value = percent_change
          
                If ws.Cells(row_tracker, 10).Value > 0 Then
                    ws.Cells(row_tracker, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(row_tracker, 10).Value < 0 Then
                    ws.Cells(row_tracker, 10).Interior.ColorIndex = 3
                End If
                
                
                If percent_change > maxValue Then
                    ticker1 = ws.Cells(row_tracker, 9).Value
                    maxValue = percent_change
                   
                
            End If
                
                
                If percent_change < minValue Then
                    ticker2 = ws.Cells(row_tracker, 9).Value
                    minValue = percent_change
                    
               
               End If
               
            
                maxVolume = WorksheetFunction.Max(ws.Range("L2:L" & rowcount))
                  ticker3 = ws.Cells(row_tracker, 9).Value
            
        
                
              

                row_tracker = row_tracker + 1
                total_volume = 0
                start = i + 1
                
            Else
                total_volume = total_volume + ws.Cells(i, 7).Value
            
            End If
        Next i
        
        ws.Range("Q2").Value = maxValue
        ws.Range("Q3").Value = minValue
        ws.Range("Q4").Value = maxVolume
        ws.Range("P2").Value = ticker1
        ws.Range("P3").Value = ticker2
        ws.Range("P4").Value = ticker3
        
    Next ws


End Sub
