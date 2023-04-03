Attribute VB_Name = "Module12"
Sub Percentchange()

  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
   
  Dim Percent_Change As Double
  Percent_Change = 2
  Dim Percentchange_Row As Integer
  Percentchange_Row = 2
  

  Dim J As Long
  J = 2
  
  
  For i = 2 To lastrow
  If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

  Percent_Change = ((Cells(i, 6).Value - Cells(J, 3).Value) / (Cells(J, 3).Value))
    
  J = i + 1

    
  Range("K" & Percentchange_Row).Value = Percent_Change
  Percentchange_Row = Percentchange_Row + 1
  
  Cells(1, 11).Value = "Percent Change"
 
  Columns("A:L").AutoFit
  
  End If
  
  
  Cells(i, 11).Value = Format(Percent_Change, "Percent")
  
  
  
  
  Next i


 







End Sub
