Sub Stock_Analysis()

    Dim WS_Count, i As Integer
    Dim last_row As Long
    Dim Ticker As String
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Total_Stock_Volume As Double
    
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    For i = 1 To WS_Count
        Sheet_Name = ActiveWorkbook.Worksheets(i).Name
        
        last_row = Worksheets(Sheet_Name).Cells(Rows.Count, 1).End(xlUp).Row
        
        'Create Column Names
        Worksheets(Sheet_Name).Range("I1").Value = "Ticker"
        Worksheets(Sheet_Name).Range("J1").Value = "Yearly Change"
        Worksheets(Sheet_Name).Range("K1").Value = "Percent Change"
        Worksheets(Sheet_Name).Range("L1").Value = "Total Stock Volume"
    
        Worksheets(Sheet_Name).Range("S1").Value = "Opening Price"
        Worksheets(Sheet_Name).Range("T1").Value = "Closing Price"

        Worksheets(Sheet_Name).Range("O2").Value = "Greatest % Increase"
        Worksheets(Sheet_Name).Range("O3").Value = "Greatest % Decrease"
        Worksheets(Sheet_Name).Range("O4").Value = "Greatest Total Volume"
        Worksheets(Sheet_Name).Range("P1").Value = "Ticker"
        Worksheets(Sheet_Name).Range("Q1").Value = "Value"
        
        'Assign values to variables
        Total_Stock_Volume = 0
        Summary_Table_Row = 2
        Summary_Table_Row2 = 2
        Summary_Table_Row3 = 2
        Max = 0
        Min = 0
        Total_Stock_Volume_Max = 0
        
        For j = 2 To last_row
            If Worksheets(Sheet_Name).Cells(j + 1, 1).Value <> Worksheets(Sheet_Name).Cells(j, 1).Value Then
                Ticker = Worksheets(Sheet_Name).Cells(j, 1).Value
                Closing_Price = Worksheets(Sheet_Name).Cells(j, 6).Value
                Total_Stock_Volume = Total_Stock_Volume + Worksheets(Sheet_Name).Cells(j, 7).Value
            
                Worksheets(Sheet_Name).Range("I" & Summary_Table_Row).Value = Ticker
                Worksheets(Sheet_Name).Range("T" & Summary_Table_Row).Value = Closing_Price
                Worksheets(Sheet_Name).Range("L" & Summary_Table_Row).Value = Total_Stock_Volume

                Summary_Table_Row = Summary_Table_Row + 1
                Total_Stock_Volume = 0

            ElseIf Worksheets(Sheet_Name).Cells(j - 1, 1).Value <> Worksheets(Sheet_Name).Cells(j, 1).Value Then
                Opening_Price = Worksheets(Sheet_Name).Cells(j, 3).Value
                Worksheets(Sheet_Name).Range("S" & Summary_Table_Row).Value = Opening_Price
        
            Else
                Total_Stock_Volume = Total_Stock_Volume + Worksheets(Sheet_Name).Cells(j, 7).Value
            End If
        Next j
        
        last_row2 = Worksheets(Sheet_Name).Cells(Rows.Count, 19).End(xlUp).Row
        
        'loop through each row of opening year and closing year stock prices for each ticker
        'calculate yearly change values
        For j = 2 To last_row2
            Yearly_Change = Worksheets(Sheet_Name).Cells(j, 20).Value - Worksheets(Sheet_Name).Cells(j, 19).Value
            Worksheets(Sheet_Name).Range("J" & Summary_Table_Row2).Value = Yearly_Change
            Summary_Table_Row2 = Summary_Table_Row2 + 1
        Next j
        
        'loop through each row of unique tickers
        'format Yearly Change values (green for positive and red for negative)
        For j = 2 To last_row2
            If Worksheets(Sheet_Name).Cells(j, 10).Value > 0 Then
                Worksheets(Sheet_Name).Cells(j, 10).Interior.ColorIndex = 4
            ElseIf Worksheets(Sheet_Name).Cells(j, 10).Value < 0 Then
                Worksheets(Sheet_Name).Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
        
        'loop through each row of unique tickers
        'create percent change values
        'Note: if the opening price was 0 then the percent change value is missing
        For j = 2 To last_row2
            If Worksheets(Sheet_Name).Cells(j, 19).Value > 0 Then
                Percent_Change = (Worksheets(Sheet_Name).Cells(j, 10).Value / Worksheets(Sheet_Name).Cells(j, 19).Value)
                Worksheets(Sheet_Name).Range("K" & Summary_Table_Row3).Value = Percent_Change
                Worksheets(Sheet_Name).Range("K" & Summary_Table_Row3).NumberFormat = "0.00%"
                Summary_Table_Row3 = Summary_Table_Row3 + 1
            End If
        Next j
        
        'loop through each row of unique tickers
        'find the tickers with the greatest increase and decrease in percent change along with their percent changes
        For j = 2 To last_row2
            If Worksheets(Sheet_Name).Cells(j, 11).Value > Max Then
                Max = Worksheets(Sheet_Name).Cells(j, 11).Value
                Ticker = Worksheets(Sheet_Name).Cells(j, 9).Value
                Worksheets(Sheet_Name).Range("Q2").Value = Max
                Worksheets(Sheet_Name).Range("Q2").NumberFormat = "0.00%"
                Worksheets(Sheet_Name).Range("P2").Value = Ticker
            End If
            If Worksheets(Sheet_Name).Cells(j, 11).Value < Min Then
                Min = Worksheets(Sheet_Name).Cells(j, 11).Value
                Ticker = Worksheets(Sheet_Name).Cells(j, 9).Value
                Worksheets(Sheet_Name).Range("Q3").Value = Min
                Worksheets(Sheet_Name).Range("Q3").NumberFormat = "0.00%"
                Worksheets(Sheet_Name).Range("P3").Value = Ticker
            End If
            If Worksheets(Sheet_Name).Cells(j, 12).Value > Total_Stock_Volume_Max Then
                Total_Stock_Volume_Max = Worksheets(Sheet_Name).Cells(j, 12).Value
                Ticker = Worksheets(Sheet_Name).Cells(j, 9).Value
                Worksheets(Sheet_Name).Range("Q4").Value = Total_Stock_Volume_Max
                Worksheets(Sheet_Name).Range("P4").Value = Ticker
            End If
        Next j
    
    Next i
End Sub