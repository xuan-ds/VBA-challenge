Attribute VB_Name = "Module1"
Sub VBA_M2_Challenge()
 
 Dim ws As Worksheet
 
For Each ws In Worksheets
    ws.Select
     Dim LR As Long
      LR = ws.Cells(1, 1).End(xlDown).Row

     Dim Ticker As String
     Dim Open_Price As Double
      Open_Price = ws.Cells(2, 3).Value
     Dim Close_Price As Double
      Close_Price = 0
     Dim Vol As LongLong
      Vol = 0
     Dim YearlyChange As Double
     Dim PercentChange As Double
     Dim TotalStockVolume As LongLong

    'Create column headers
     ws.Cells(1, 9).Value = "Ticker"
     ws.Cells(1, 10).Value = "Yearly Change"
     ws.Cells(1, 11).Value = "Percentage Change"
     ws.Cells(1, 12).Value = "Total Stock Volume"
     ws.Cells(2, 15).Value = "Greatest % Increase"
     ws.Cells(3, 15).Value = "Greatest % Decrease"
     ws.Cells(4, 15).Value = "Greatest Total Volume"
   
    Dim i As Long
    Dim j As Integer
           j = 2
   
       For i = 2 To LR
   
        If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
                 Vol = Vol + ws.Cells(i, 7).Value
              
        Else
             YearlyChange = ws.Cells(i, 6).Value - Open_Price
                 ws.Cells(j, 10).Value = YearlyChange
                 
                 If ws.Cells(j, 10).Value >= 0 Then
                   ws.Cells(j, 10).Interior.ColorIndex = 4
                  Else
                   ws.Cells(j, 10).Interior.ColorIndex = 3
                 End If
            
            PercentChange = YearlyChange / Open_Price
                 ws.Cells(j, 11).Value = PercentChange
                 ws.Cells(j, 11).NumberFormat = "0.00%"
    
            TotalStockVolume = Vol + ws.Cells(i, 7).Value
                 ws.Cells(j, 9).Value = ws.Cells(i, 1)
               
                  ws.Cells(j, 12).Value = TotalStockVolume
               
                 Open_Price = ws.Cells(i + 1, 3).Value
                Vol = 0
               j = j + 1
        
            End If
    
  
        Next i
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Prepare for summary
        GreatVol = ws.Cells(2, 12).Value
        GreatIncr = ws.Cells(2, 11).Value
        GreatDecr = ws.Cells(2, 11).Value
        
            For i = 2 To LastRowI
            
                If ws.Cells(i, 12).Value > GreatVol Then
                GreatVol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatVol = GreatVol
                
                End If
                
                If ws.Cells(i, 11).Value > GreatIncr Then
                GreatIncr = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatIncr = GreatIncr
                
                End If
            
                If ws.Cells(i, 11).Value < GreatDecr Then
                GreatDecr = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatDecr = GreatDecr
                
                End If
                
            ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
            
            Next i
     
     Next ws
    
    
End Sub

