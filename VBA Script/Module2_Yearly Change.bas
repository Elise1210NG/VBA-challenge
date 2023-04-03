Attribute VB_Name = "Module11"
Sub YearlyChange()

  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
  Dim Ticker_open As Double
  Ticker_open = 0
  Dim Ticker_close As Double
  Ticker_close = 0
  Dim Yearly_change As Double
  Yearly_change = 2
  Dim Pricechange_Row As Integer
  Pricechange_Row = 2
  
  
  Dim J As Long
  J = 2
  
  
  For i = 2 To lastrow
  If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
   
  Ticker_close = Cells(i, 6).Value
  Yearly_change = Cells(i, 6).Value - Cells(J, 3).Value
  
  
  J = i + 1
     
  Range("J" & Pricechange_Row).Value = Yearly_change
  Pricechange_Row = Pricechange_Row + 1
    
  
  Cells(1, 10).Value = "Yearly change"

  
  Columns("A:L").AutoFit
  
  End If
  
  If Cells(i, 10).Value < 0 Then
  Cells(i, 10).Interior.ColorIndex = 3
  
  ElseIf Cells(i, 10).Value >= 0 Then
  Cells(i, 10).Interior.ColorIndex = 4
  

  
  End If

  
  
  
  Next i


 







End Sub
