Sub stockDataPerYear()

For Each ws In Worksheets
ws.Activate

Dim Worksheet As String


Worksheet = ws.Name
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  Dim ticker As String
  
  Dim InitialOpening As Double
  InitialOpening = Cells(2, 3).Value
  
  Dim initialrow As Integer
  initialrow = 2
  
  Dim Initialclosing As Double
  Initialclosing = Cells(2, 3).Value
  
  Dim closingrow As Integer
  closingrow = 2
  
  Dim Yearlychangerow As Integer
  Yearlychangerow = 2
  
  Dim Yearlychange As Double
  Yearlychange = 0
 
  Dim Percentchange As Double
  Percentchange = 0
  
  Dim totalvol As Double
  totalvol = 0

  Dim TotalVolrow As Integer
  TotalVolrow = 2
  
    For i = 2 To lastrow
    
    
  
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ticker = Cells(i, 1).Value

      totalvol = totalvol + Cells(i, 3).Value
    
      Range("I" & TotalVolrow).Value = ticker

      Range("L" & TotalVolrow).Value = totalvol
      
      Yearlychange = Cells(i, 6).Value - InitialOpening
    
      
      Range("J" & TotalVolrow).Value = Yearlychange
      
      
      
      If InitialOpening <> 0 Then
        Percentchange = (Yearlychange / InitialOpening)
        Else
          Percentchange = 0
      End If
      Range("K" & TotalVolrow).NumberFormat = "0.00%"
      
      Range("K" & TotalVolrow).Value = Percentchange
      
      If Percentchange < 0 Then
      Range("K" & TotalVolrow).Interior.ColorIndex = 3
      
      Else
      Range("K" & TotalVolrow).Interior.ColorIndex = 4
      
      End If
      
      

      TotalVolrow = TotalVolrow + 1
    
      
      totalvol = 0
      InitialOpening = Cells(i + 1, 3)
      
      If Cells(i + 1, 3) = 0 Then
        InitialOpening = Cells(i + 1, 6)
        
    End If
        
    Else

      totalvol = totalvol + Cells(i, 3).Value
      

    End If
    
    
    

  Next i
  Dim GreatestIncrease As Double
  Dim Greatestdecrease As Double
  Dim Greatesttotalvolume As Double
  
  
  
  GreatestIncrease = Cells(2, 11).Value
  Greatestdecrease = Cells(2, 11).Value
  Greatesttotalvol = Cells(2, 12).Value
  
  
  lastrow = ws.Cells(Rows.Count, 11).End(xlUp).Row
  
  For i = 2 To lastrow
  
    If Cells(i, 11).Value > GreatestIncrease Then
    
        ticker = Cells(i, 9).Value
    
        GreatestIncrease = Cells(i, 11).Value
        
    End If
    
    If Cells(i, 11).Value < Greatestdecrease Then
    
     ticker1 = Cells(i, 9).Value
     Greatestdecrease = Cells(i, 11).Value
     
     End If
     
     If Cells(i, 12).Value > Greatesttotalvol Then
     
     ticker2 = Cells(i, 9).Value
     Greatesttotalvol = Cells(i, 12).Value
     
    End If
    
   
    
     
    
  
 Next i
 
 Cells(2, 16).Value = GreatestIncrease
 
 Cells(2, 15).Value = ticker
 
 Cells(3, 16).Value = Greatestdecrease
 
 Cells(3, 15).Value = ticker1
 
 Cells(4, 15).Value = ticker2
 
 Cells(4, 16).Value = Greatesttotalvol
 
 
 
Next ws
  
  Exit Sub

End Sub
