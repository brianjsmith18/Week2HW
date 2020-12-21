Attribute VB_Name = "Module1"
Sub stock_data()

Dim Stock_Ticker As String
Dim Yearly_Change As Double
Yearly_Change = 0
Dim Percent_Change As Double
Percent_Change = 0
Dim Total_Stock_Volume As LongLong
Total_Stock_Volume = 0
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
Dim Open_Price As Double
Open_Price = Cells(2, 3).Value
Dim Close_Price As Double
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Volume"

For i = 2 To 70926
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Stock_Ticker = Cells(i, 1).Value
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7)
        Range("I" & Summary_Table_Row).Value = Stock_Ticker
        Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
        Close_Price = Cells(i, 6).Value
        Change = Close_Price - Open_Price
        Open_Price = Cells(i + 1, 3).Value
        Range("J" & Summary_Table_Row).Value = Change
        Range("J" & Summary_Table_Row).NumberFormat = "0.00"
             If Range("J" & Summary_Table_Row).Value >= 0 Then
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            ElseIf Range("J" & Summary_Table_Row).Value <= 0 Then
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            
            End If
        Percent_Change = Change / Cells(2, 3).Value
        Range("K" & Summary_Table_Row).Value = Percent_Change
        Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
        
        Summary_Table_Row = Summary_Table_Row + 1
        Total_Stock_Volume = 0
        
    Else
    
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        
    End If

Next i

End Sub

