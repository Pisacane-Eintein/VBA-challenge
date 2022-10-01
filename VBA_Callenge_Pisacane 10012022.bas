Attribute VB_Name = "Module1"
Option Explicit

Sub Stock_Table()

Dim Stock_Name As String
Dim Stock_Volume As Double
Dim Price_Change_Amount As Double
Dim Price_Change_Percent As Double
Dim Beg_Price As Double
Dim End_Price As Double
Dim i As Double

Stock_Volume = 0
Beg_Price = 0
End_Price = 0


Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Annual Stock Price Change"
Cells(1, 11).Value = "Percentage Annual Stock Price Change"
Cells(1, 12).Value = "Total Stock Volume"

Dim Summary_Row As Integer
Summary_Row = 2

    For i = 2 To 753001
        
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            Beg_Price = Cells(i, 3).Value
        End If
        
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            
            Stock_Name = Cells(i, 1)
            Stock_Volume = Stock_Volume + Cells(i, 7).Value
            End_Price = Cells(i, 6).Value
                        
            Range("I" & Summary_Row).Value = Stock_Name
            Range("L" & Summary_Row).Value = Stock_Volume
            Range("J" & Summary_Row).Value = Beg_Price - End_Price
            Range("K" & Summary_Row).Value = (((End_Price - Beg_Price) / Beg_Price))
                
                
            If Cells(Summary_Row, 11).Value >= 0 Then
                Cells(Summary_Row, 11).SInterior.ColorIndex = 4
            Else
                Cells(Summary_Row, 11).Interior.ColorIndex = 3
          
        End If
                
        Summary_Row = Summary_Row + 1
        Stock_Volume = 0
        
        
        Else
        
        Stock_Volume = Stock_Volume + Cells(i, 7).Value
                
              
        End If
               
        Next i
        
        Range("K1").EntireColumn.NumberFormat = "#0.00%"
        Range("J1").EntireColumn.NumberFormat = "#0.00"
      
        
End Sub

    
    




