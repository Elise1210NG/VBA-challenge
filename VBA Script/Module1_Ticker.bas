Attribute VB_Name = "Module1"
Sub Ticker()

  
  Dim Ticker As String
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
    
  
  For i = 2 To lastrow
  If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
  Ticker = Cells(i, 1).Value
  
  
  Range("I" & Summary_Table_Row).Value = Ticker
  Summary_Table_Row = Summary_Table_Row + 1
  
  
 
        
  Cells(1, 9).Value = "Ticker"
  
  Columns("A:L").AutoFit
  
  End If
  
 
  Next i


 
End Sub
