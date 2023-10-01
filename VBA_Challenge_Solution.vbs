Attribute VB_Name = "Module1"

Option Explicit
Sub Module2Challenge():
  
    Const FIRST_DATA_ROW As Integer = 2
    Const IN_TICKER_COL As Integer = 1
    Const OPEN_COL As Integer = 3
    Const CLOSE_COL As Integer = 6

    
    
    
    
    Dim ws As Worksheet
    Dim prev_ticker As String
    Dim current_ticker As String
    Dim next_ticker As String
    Dim input_row As Long
    Dim last_data_row As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_vol As Double
    Dim greatest_increase As Double
    Dim greatest_increase_ticker As String
    Dim greatest_decrease_ticker As String
    Dim greatest_total_vol_ticker As String
    Dim greatest_decrease As Double
    Dim greatest_total_vol As Double
    Dim first As Boolean
    Dim current_summary_row As Integer

    For Each ws In Worksheets
        ws.Activate
        last_data_row = Cells(FIRST_DATA_ROW, IN_TICKER_COL).End(xlDown).Row
        current_summary_row = 2
        first = True
        total_vol = 0
        For input_row = FIRST_DATA_ROW To last_data_row
            prev_ticker = Cells(input_row - 1, IN_TICKER_COL).Value
            current_ticker = Cells(input_row, IN_TICKER_COL).Value
            next_ticker = Cells(input_row + 1, IN_TICKER_COL).Value
            total_vol = total_vol + Cells(input_row, 7).Value
            'First row of stock
            If current_ticker <> prev_ticker Then
                open_price = Cells(input_row, OPEN_COL).Value
            End If
            
            
            'Last row of stock
            If current_ticker <> next_ticker Then
            
                close_price = Cells(input_row, CLOSE_COL).Value
                Range("I" & current_summary_row).Value = current_ticker
                
                yearly_change = close_price - open_price
                Range("J" & current_summary_row).Value = yearly_change
                Range("K" & current_summary_row).Value = (yearly_change / open_price)
                If first = True Then
                    greatest_increase = (yearly_change / open_price)
                    greatest_decrease = (yearly_change / open_price)
                    greatest_total_vol = total_vol
                    greatest_increase_ticker = current_ticker
                    greatest_decrease_ticker = current_ticker
                    greatest_total_vol_ticker = current_ticker
                    first = False
                End If
            
                If (yearly_change / open_price) > greatest_increase Then
                    greatest_increase = (yearly_change / open_price)
                    greatest_increase_ticker = current_ticker
                End If
                
                If (yearly_change / open_price) < greatest_decrease Then
                    greatest_decrease = (yearly_change / open_price)
                    greatest_decrease_ticker = current_ticker
                End If
                
                If (total_vol > greatest_total_vol) Then
                    greatest_total_vol = total_vol
                    greatest_total_vol_ticker = current_ticker
                End If
                
                If yearly_change < 0 Then
                    Range("J" & current_summary_row).Interior.Color = RGB(255, 0, 0)
                ElseIf yearly_change > 0 Then
                    Range("J" & current_summary_row).Interior.Color = RGB(0, 128, 0)
                End If
                
                Range("K" & current_summary_row).Style = "Percent"
                Range("L" & current_summary_row).Value = total_vol
                
                current_summary_row = current_summary_row + 1
                total_vol = 0
            End If
            
        Next input_row
        
        Range("Q" & 2).Value = greatest_increase
        Range("Q" & 2).Style = "Percent"
        Range("P" & 2).Value = greatest_increase_ticker
        Range("Q" & 3).Value = greatest_decrease
         Range("Q" & 3).Style = "Percent"
        Range("P" & 3).Value = greatest_decrease_ticker
        Range("Q" & 4).Value = greatest_total_vol
        Range("P" & 4).Value = greatest_total_vol_ticker
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
    Next ws
    
    MsgBox ("done")
End Sub


