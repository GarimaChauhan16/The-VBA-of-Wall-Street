Attribute VB_Name = "Module1"
Sub Stocks()
For Each ws In Worksheets
    Dim last_row As Long
    Dim Total_Volume, new_row As Long
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    Total_Volume = 0
    new_row = 2
    
'Part1: Ticker Symbol
    For i = 2 To last_row
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
        Stock_Name = ws.Cells(i, 1).Value
        ws.Range("I" & new_row).Value = Stock_Name
        new_row = new_row + 1
        End If
    Next i
    
'Part2: Yearly Price Change
    new_row = 2
    first_day = 2
    For i = 2 To last_row
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
        last_day = i
        first_day_open_price = ws.Range("C" & first_day)
        last_day_close_price = ws.Range("F" & last_day)
        yearly_change = last_day_close_price - first_day_open_price
        ws.Range("J" & new_row).Value = yearly_change
        first_day = i + 1
        
'Part3: Percent Change
            If first_day_open_price = 0 And last_day_close_price = 0 Then
            percent_change = 0
            ElseIf first_day_open_price = 0 And last_day_close_price <> 0 Then
            percent_change = 1
            Else
            percent_change = yearly_change / first_day_open_price
            ws.Range("K" & new_row).Value = percent_change
            ws.Range("K" & new_row).NumberFormat = "0.00%"
            End If
'conditional Formatting
            If yearly_change >= 0 Then
            ws.Range("J" & new_row).Interior.ColorIndex = 4
            Else
            ws.Range("J" & new_row).Interior.ColorIndex = 3
            End If
        new_row = new_row + 1
        End If
    Next i
        
'Part4: Total Volume Calculations
    new_row = 2
    For i = 2 To last_row
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
        Total_Volume = Total_Volume + ws.Range("G" & i).Value
        ws.Range("L" & new_row).Value = Total_Volume
        Total_Volume = 0
        new_row = new_row + 1
        Else
        Total_Volume = Total_Volume + ws.Range("G" & i).Value
        End If
    Next i
    
'Part5: Greatest Value Calculations
    ws.Range("O3").Value = "Greatest % Increase"
    ws.Range("O4").Value = "Greatest % Decrease"
    ws.Range("O5").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    Greatest_Increase = ws.Range("K2").Value
    Greatest_Decrease = ws.Range("K2").Value
    Greatest_Total_Volume = ws.Range("L2").Value
'Greatest % Increase
    For i = 2 To new_row
        If ws.Range("K" & i).Value >= Greatest_Increase Then
        Greatest_Increase = ws.Range("K" & i).Value
        ws.Range("P3").Value = ws.Range("I" & i)
        ws.Range("Q3").Value = Greatest_Increase
        ws.Range("Q3").NumberFormat = "0.00%"
        End If
'Greatest % Decrease
        If ws.Range("K" & i).Value <= Greatest_Decrease Then
        Greatest_Decrease = ws.Range("K" & i).Value
        ws.Range("P4").Value = ws.Range("I" & i)
        ws.Range("Q4").Value = Greatest_Decrease
        ws.Range("Q4").NumberFormat = "0.00%"
        End If
'Greatest Total Volume
        If ws.Range("L" & i).Value >= Greatest_Total_Volume Then
        Greatest_Total_Volume = ws.Range("L" & i).Value
        ws.Range("P5").Value = ws.Range("I" & i)
        ws.Range("Q5").Value = Greatest_Total_Volume
        End If
    Next i
    ws.Columns("I:Q").EntireColumn.AutoFit
Next ws
End Sub
