Sub Ticker():

  Dim Last As Long
  Dim Counter As Integer
  Dim Ticker As String
  Dim V1 As Double
  Dim V2 As Double
  Dim Volume As Double
  Dim GInc As Double
  Dim GDec As Double
  Dim GVol As Double
  Dim IncTick As String
  Dim DecTick As String
  Dim VolTick As String
  
  Volume = 0
  Counter = 2
  GInc = 0
  GDec = 0
  GVol = 0
  
  Cells(1, 9).Value = "Ticker"
  Cells(1, 10).Value = "Yearly Change"
  Cells(1, 11).Value = "Percentage Change"
  Cells(1, 12).Value = "Total Stock Volume"
  Cells(1, 16).Value = "Ticker"
  Cells(1, 17).Value = "Value"
  
  Cells(2, 15).Value = "Greatest % Increase"
  Cells(3, 15).Value = "Greatest % Decrease"
  Cells(4, 15).Value = "Greatest Total Volume"
  
  For x = 1 To 1000000:
    
    If Cells(x + 1, 1).Value = "" Then
      
      Last = x
      
      Exit For
    
    End If
    
  Next x
  
  For i = 2 To Last:
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      
      Ticker = Cells(i, 1).Value
      V2 = Cells(i, 6).Value
      Volume = Volume + Cells(i, 7).Value
      
      For Z = i To 1 Step -1:
        
        If Cells(Z - 1, 1).Value <> Ticker Then
          
          V1 = Cells(Z, 3).Value
          
          Exit For
          
        End If
        
      Next Z
      
      Cells(Counter, 9).Value = Ticker
      Cells(Counter, 10).Value = V2 - V1
      Cells(Counter, 11).Value = FormatPercent((Cells(Counter, 10)) / V1)
      Cells(Counter, 12).Value = Volume
      
      If Cells(Counter, 10).Value >= 0 Then
      
        Cells(Counter, 10).Interior.ColorIndex = 4
      
      Else:
        
        Cells(Counter, 10).Interior.ColorIndex = 3
        
      End If
      
      Counter = Counter + 1
      Volume = 0
    
    Else:
      
      Volume = Volume + Cells(i, 7).Value
      
    End If
    
  Next i
  
  For a = 2 To Last:
  
    If Cells(a, 11) > GInc Then
      GInc = Cells(a, 11)
      IncTick = Cells(a, 9)
    End If
    
    If Cells(a, 11) < GDec Then
      GDec = Cells(a, 11)
      DecTick = Cells(a, 9)
    End If
    
    If Cells(a, 12).Value > GVol Then
      GVol = Cells(a, 12).Value
      VolTick = Cells(a, 9)
    End If
  
  Next a
  
  Cells(2, 16).Value = IncTick
  Cells(2, 17).Value = FormatPercent(GInc)
  Cells(3, 16).Value = DecTick
  Cells(3, 17).Value = FormatPercent(GDec)
  Cells(4, 16).Value = VolTick
  Cells(4, 17).Value = GVol

End Sub