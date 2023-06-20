Attribute VB_Name = "Module1"


Sub stock_analysis()
    Dim last_row As Long
    Dim column As Long
    Dim closing_price As Double
    Dim opening_price As Double
    Dim yearly_change As Double
    Dim percentage_change As Double
    Dim increment As Long
    Dim max_increase_ticker As String
    Dim max_increase_value As Double
    Dim max_decrease_ticker As String
    Dim max_decrease_value As Double
    Dim max_volume_ticker As String
    Dim max_volume_value As Double
    
    increment = 0
    column = 1
    max_increase_value = 0
    max_decrease_value = 0
    max_volume_value = 0

    ' Find the last row of the active sheet
    last_row = Cells(Rows.count, 1).End(xlUp).Row

    ' Loop through all the rows
    For i = 2 To last_row
        ' When it finds a different value in the ticker row
        If Cells(i + 1, column).Value <> Cells(i, column).Value Then
            increment = increment + 1
            ' Return ticker
            Cells(1 + increment, 9) = Cells(i, 1).Value
            ' Return the difference between the end year closing price and the beginning of the year open price
            closing_price = Cells(i, 6).Value
            opening_price = Cells(i - 250, 3).Value
            yearly_change = closing_price - opening_price
            Cells(1 + increment, 10).Value = yearly_change
            ' Show percentage change in price
            percentage_change = (yearly_change / opening_price) * 100
            Cells(1 + increment, 11).Value = percentage_change
            Cells(1 + increment, 11).NumberFormat = "0.00%"
            ' Show total stock volume
            Cells(1 + increment, 12).Formula = "=SUM(" & Range(Cells(i - 250, 7), Cells(i, 7)).Address(False, False) & ")"
            
            If yearly_change > 0 Then
                Cells(1 + increment, 10).NumberFormat = "0.00"
                Cells(1 + increment, 10).Interior.ColorIndex = 4 ' Format positive change in green
            ElseIf yearly_change < 0 Then
                Cells(1 + increment, 10).NumberFormat = "0.00"
                Cells(1 + increment, 10).Interior.ColorIndex = 3 ' Format negative change in red
            Else
                Cells(1 + increment, 10).NumberFormat = "0.00"
                Cells(1 + increment, 10).Interior.ColorIndex = xlNone ' Clear interior color for no change
            End If
            
            ' Update maximum % increase
            If percentage_change > max_increase_value Then
                max_increase_value = percentage_change
                max_increase_ticker = Cells(i, 1).Value
            End If
            
            ' Update maximum % decrease
            If percentage_change < max_decrease_value Then
                max_decrease_value = percentage_change
                max_decrease_ticker = Cells(i, 1).Value
            End If
            
            ' Update maximum total volume
            If Cells(1 + increment, 12).Value > max_volume_value Then
                max_volume_value = Cells(1 + increment, 12).Value
                max_volume_ticker = Cells(i, 1).Value
            End If
        End If
    Next i
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percentage Change"
    Range("L1").Value = "Total Stock Volume"
    
    ' Generate small table
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    Range("P2").Value = max_increase_ticker
    Range("Q2").Value = max_increase_value
    Range("Q2").NumberFormat = "0.00%"
    
    Range("P3").Value = max_decrease_ticker
    Range("Q3").Value = max_decrease_value
    Range("Q3").NumberFormat = "0.00%"
    
    Range("P4").Value = max_volume_ticker
    Range("Q4").Value = max_volume_value
    
    ' Auto-fit columns
    Columns("I:L").AutoFit
    Columns("O:Q").AutoFit
End Sub

