Attribute VB_Name = "Module3"
Sub Combine()

  Dim Ticker As String
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  lastrow2 = Cells(Rows.Count, 9).End(xlUp).Row
  
  Dim Ticker_open As Double
  Ticker_open = 0
  Dim Ticker_close As Double
  Ticker_close = 0
  Dim Yearly_change As Double
  Yearly_change = 2
  Dim Pricechange_Row As Integer
  Pricechange_Row = 2
     
  Dim Percent_Change As Double
  Percent_Change = 2
  Dim Percentchange_Row As Integer
  Percentchange_Row = 2
  
  Dim Total_volume As Double
  Total_volume = 0
  Dim Totalvolume_Row As Integer
  Totalvolume_Row = 2
  
  Dim Greatperce_inc As Double
  Dim Greatperce_dec As Double
  Dim Greattotal_vol As Double
  Dim ticker2 As Double
  
  Dim J As Long
  J = 2
    
  For i = 2 To lastrow
  If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
  Ticker = Cells(i, 1).Value
  
  
  Range("I" & Summary_Table_Row).Value = Ticker
  Summary_Table_Row = Summary_Table_Row + 1
    
  Ticker_close = Cells(i, 6).Value
  Yearly_change = Cells(i, 6).Value - Cells(J, 3).Value
  Percent_Change = ((Cells(i, 6).Value - Cells(J, 3).Value) / (Cells(J, 3).Value))
  Total_volume = WorksheetFunction.Sum(Range(Cells(J, 7), Cells(i, 7)))
  
  J = i + 1
     
  Range("J" & Pricechange_Row).Value = Yearly_change
  Pricechange_Row = Pricechange_Row + 1
    
  Range("K" & Percentchange_Row).Value = Percent_Change
  Percentchange_Row = Percentchange_Row + 1
  
  Range("L" & Totalvolume_Row).Value = Total_volume
  Totalvolume_Row = Totalvolume_Row + 1
        
  Cells(1, 9).Value = "Ticker"
  Cells(1, 10).Value = "Yearly change"
  Cells(1, 11).Value = "Percent Change"
  Cells(1, 12).Value = "Total Stock Volume"
  Cells(2, 15).Value = "Greatest % Increase"
  Cells(3, 15).Value = "Greatest % Decrease"
  Cells(4, 15).Value = "Greatest Total Volume"
  Cells(1, 16).Value = "Ticker"
  Cells(1, 17).Value = "Value"
  
  Columns("A:L").AutoFit
  
  End If
  
  If Cells(i, 10).Value >= 0 Then
  Cells(i, 10).Interior.ColorIndex = 4
  
  
  ElseIf Cells(i, 10).Value < 0 Then
  Cells(i, 10).Interior.ColorIndex = 3
  
  End If
  
  Cells(i, 11).Value = Format(Percent_Change, "Percent")
  
  Next i
  
  
  For i = 2 To lastrow2
  
  Greatperce_inc = WorksheetFunction.Max(Range(Cells(2, 11), Cells(i, 11)))
  Cells(2, 17).Value = Greatperce_inc
  Cells(2, 17).Value = Format(Greatperce_inc, "Percent")
  
 
  
  Greatperce_dec = WorksheetFunction.Min(Range(Cells(2, 11), Cells(i, 11)))
  Cells(3, 17).Value = Greatperce_dec
  Cells(3, 17).Value = Format(Greatperce_dec, "Percent")
  
  
  
  Greattotal_vol = WorksheetFunction.Max(Range(Cells(2, 12), Cells(i, 12)))
  Cells(4, 17).Value = Greattotal_vol
  
   
  
  
  Columns("O:Q").AutoFit
   
  
  Next i
  


End Sub
