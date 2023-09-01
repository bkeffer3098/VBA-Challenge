Sub Stock_Data()

For Each Worksheet In ThisWorkbook.Sheets
    Worksheet.Activate

Dim Ticker, Ticker_Increase, Ticker_Decrease, Ticker_Volume As String

Dim Yearly_Change, Percent_Change, open_price, close_price, Percent_Increase, Percent_Decrease, Greatest_Volume As Double
Percent_Increase = 0
Percent_Decrease = 0
Greatest_Volume = 0
    
Dim Volume As Double
    Volume = 0

Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

Dim lastrow As Long
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row

'Dim Ticker_Title, Yearly, Percent, Total As String
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

open_price = Cells(2, 3).Value

    For i = 2 To lastrow
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            Ticker = Cells(i, 1).Value
            
            close_price = Cells(i, 6).Value
            
            Yearly_Change = close_price - open_price
            Cells(i, 10).Value = Yearly_Change
            
            Percent_Change = (close_price - open_price) / open_price
            
            open_price = Cells(i + 1, 3).Value
            
            Volume = Volume + Cells(i, 7).Value
            
            Range("I" & Summary_Table_Row).Value = Ticker
            
            Range("J" & Summary_Table_Row).Value = Yearly_Change
            
            Range("L" & Summary_Table_Row).Value = Volume
            
                If Greatest_Volume < Volume Then
        
                    Greatest_Volume = Volume
                    Ticker_Volume = Ticker
            
                End If
            
            Range("K" & Summary_Table_Row).Value = Percent_Change
                
                If Percent_Increase < Percent_Change Then
        
                    Percent_Increase = Percent_Change
                    Ticker_Increase = Ticker
            
                End If
                
                If Percent_Decrease > Percent_Change Then
        
                    Percent_Decrease = Percent_Change
                    Ticker_Decrease = Ticker
            
                End If
            
            Summary_Table_Row = Summary_Table_Row + 1
            
            Volume = 0
            
        Else
        
            Volume = Volume + Cells(i, 7).Value
        
        End If
    
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        Range("K" & i).NumberFormat = "0.00%"
        Range("Q2").NumberFormat = "0.00%"
        Range("Q3").NumberFormat = "0.00%"
        
           'Begin Percent Change Formatting

                If Cells(i, 10).Value >= 0.01 Then
        
                    Cells(i, 10).Interior.ColorIndex = 4
        
                ElseIf Cells(i, 10).Value < 0.01 Then
        
                    Cells(i, 10).Interior.ColorIndex = 3
    
                End If
    
    Next i
    
Range("Q2").Value = Percent_Increase
Range("Q3").Value = Percent_Decrease
Range("Q4").Value = Greatest_Volume
Range("P2").Value = Ticker_Increase
Range("P3").Value = Ticker_Decrease
Range("P4").Value = Ticker_Volume

Range("A:Q").Columns.AutoFit

Next Worksheet

End Sub



