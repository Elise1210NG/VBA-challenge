Attribute VB_Name = "Module2"
Sub totalvolnbonus()

  Dim Total_volume As Double
  Total_volume = 0
  Dim Totalvolume_Row As Integer
  Totalvolume_Row = 2
  
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
  Dim J As Long
  J = 2
  
  
  For i = 2 To lastrow
  If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
  
    
  Total_volume = WorksheetFunction.Sum(Range(Cells(J, 7), Cells(i, 7)))
  
  J = i + 1
  
  Range("L" & Totalvolume_Row).Value = Total_volume
  Totalvolume_Row = Totalvolume_Row + 1
  
  End If
  
  Next i

  lastrow2 = Cells(Rows.Count, 9).End(xlUp).Row
  
  
  Cells(1, 12).Value = "Total Stock Volume"
  Cells(2, 15).Value = "Greatest % Increase"
  Cells(3, 15).Value = "Greatest % Decrease"
  Cells(4, 15).Value = "Greatest Total Volume"
  Cells(1, 16).Value = "Ticker"
  Cells(1, 17).Value = "Value"
 
  
  Dim Greatperce_inc As Double
  Dim Greatperce_dec As Double
  Dim Greattotal_vol As Double
  Dim ticker2 As Double
  
   
   
  For i = 2 To lastrow2
  
  Greatperce_inc = WorksheetFunction.Max(Range(Cells(2, 11), Cells(i, 11)))
  Cells(2, 17).Value = Greatperce_inc
  
  Cells(2, 17).Value = Format(Greatperce_inc, "Percent")
  
  
  Greatperce_dec = WorksheetFunction.Min(Range(Cells(2, 11), Cells(i, 11)))
  Cells(3, 17).Value = Greatperce_dec
 
  Cells(3, 17).Value = Format(Greatperce_dec, "Percent")
  
    
  Greattotal_vol = WorksheetFunction.Max(Range(Cells(2, 12), Cells(i, 12)))
  Cells(4, 17).Value = Greattotal_vol
  
  Columns("L:Q").AutoFit
  
  
  Next i
 





End Sub

