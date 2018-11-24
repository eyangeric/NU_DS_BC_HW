Sub Stock_Analysis()
    
    'Create Column Names
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    Range("S1").Value = "Opening Price"
    Range("T1").Value = "Closing Price"

    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    'Create variables
    Dim Ticker As String
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Total_Stock_Volume As Double
    
    'Assign values to variables
    Total_Stock_Volume = 0
    Summary_Table_Row = 2
    Summary_Table_Row2 = 2
    Summary_Table_Row3 = 2
    Max = 0
    Min = 0
    Total_Stock_Volume_Max = 0
    
    'Loop through each row to get unique Ticker, Opening and Closing Prices for those tickers
    For i = 2 To 797711
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker = Cells(i, 1).Value
            Closing_Price = Cells(i, 6).Value
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            
            Range("I" & Summary_Table_Row).Value = Ticker
            Range("T" & Summary_Table_Row).Value = Closing_Price
            Range("L" & Summary_Table_Row).Value = Total_Stock_Volume

            Summary_Table_Row = Summary_Table_Row + 1
            Total_Stock_Volume = 0

        ElseIf Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            Opening_Price = Cells(i, 3).Value
            Range("S" & Summary_Table_Row).Value = Opening_Price
        
        Else
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        End If
    Next i

    'loop through each row of opening year and closing year stock prices for each ticker
    'calculate yearly change values
    For i = 2 To 3169
        Yearly_Change = Cells(i, 20).Value - Cells(i, 19).Value
        Range("J" & Summary_Table_Row2).Value = Yearly_Change
        Summary_Table_Row2 = Summary_Table_Row2 + 1
    Next i
    
    'loop through each row of unique tickers
    'format Yearly Change values (green for positive and red for negative)
    For i = 2 To 3169
        If Cells(i, 10).Value > 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        ElseIf Cells(i, 10).Value < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
    
    'loop through each row of unique tickers
    'create percent change values
    'Note: if the opening price was 0 then the percent change value is missing
    For i = 2 To 3169
        If Cells(i, 19).Value > 0 Then
            Percent_Change = (Cells(i, 10).Value / Cells(i, 19).Value)
            Range("K" & Summary_Table_Row3).Value = Percent_Change
            Range("K" & Summary_Table_Row3).NumberFormat = "0.00%"
            Summary_Table_Row3 = Summary_Table_Row3 + 1
        End If
    Next i

    'loop through each row of unique tickers
    'find the tickers with the greatest increase and decrease in percent change along with their percent changes
    For i = 2 To 3169
        If Cells(i, 11).Value > Max Then
            Max = Cells(i, 11).Value
            Ticker = Cells(i, 9).Value
            Range("Q2").Value = Max
            Range("Q2").NumberFormat = "0.00%"
            Range("P2").Value = Ticker
        End If
        If Cells(i, 11).Value < Min Then
            Min = Cells(i, 11).Value
            Ticker = Cells(i, 9).Value
            Range("Q3").Value = Min
            Range("Q3").NumberFormat = "0.00%"
            Range("P3").Value = Ticker
        End If
        If Cells(i, 12).Value > Total_Stock_Volume_Max Then
            Total_Stock_Volume_Max = Cells(i, 12).Value
            Ticker = Cells(i, 9).Value
            Range("Q4").Value = Total_Stock_Volume_Max
            Range("P4").Value = Ticker
        End If
    Next i
    
End Sub