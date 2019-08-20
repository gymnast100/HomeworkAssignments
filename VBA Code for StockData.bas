Attribute VB_Name = "Module1"
Sub StockData()

For Each ws In Worksheets


        Dim i As Long
        Dim TotalStock As LongLong
        Dim RowCounter As Long
         Dim LastRow As Long
           'Stock closes yearly at cells(i, 6)
           'YearlyChange = Stock closing yearly - stock opening
        Dim YearlyChange As Double
        Dim PercentageChange As Double
        Dim Stockopens As Double
        
         RowCounter = 2
        TotalStock = 0
        YearlyChange = 0
        PercentageChange = 0
        Stockopens = 0
           
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            

                        ws.Cells(1, 9).Value = "Ticker"
                        ws.Cells(1, 12).Value = " Total Stock Volume"
                        ws.Cells(1, 11).Value = " Percentage Change"
                        ws.Cells(1, 10).Value = " Yearly Change"
                        ws.Cells(2, 15).Value = " Greatest % Increase"
                        ws.Cells(3, 15).Value = " Greatest % Decrease"
                        ws.Cells(4, 15).Value = " Greatest Total Volume"
                        ws.Cells(1, 16).Value = " Ticker"
                        ws.Cells(1, 17).Value = " Value"
                        

        
            For i = 2 To LastRow

                   If Stockopens = 0 And ws.Cells(i, 3).Value <> 0 Then
                           Stockopens = ws.Cells(i, 3).Value
                   End If
            
                   If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
                   
                        YearlyChange = (ws.Cells(i, 6).Value) - Stockopens
                        ws.Cells(RowCounter, 10).Value = YearlyChange
                        PercentageChange = (YearlyChange / Stockopens)
                                    
                                    
                        ws.Cells(RowCounter, 11).Value = PercentageChange
                        ws.Cells(RowCounter, 11).Value = Format(ws.Cells(RowCounter, 11).Value, "Percent")
                                    
                        TotalStock = TotalStock + ws.Cells(i, 7).Value
                                     
                        ws.Cells(RowCounter, 9).Value = ws.Cells(i, 1)
                        ws.Cells(RowCounter, 12).Value = TotalStock
                                
                        RowCounter = RowCounter + 1
                                
        
                        TotalStock = 0
                        YearlyChange = 0
                        PercentChange = 0
                        Stockopens = ws.Cells(i + 1, 3).Value
                                   
                         Else:
                        
                            TotalStock = TotalStock + ws.Cells(i, 7).Value
                
                    
                    End If
                    
                    
            
        Next i
        
        'Color Formatting
        
        Dim LastRowYearlyChange As Long
        LastRowYearlyChange = ws.Cells(Rows.Count, 10).End(xlUp).Row
        For i = 2 To LastRowYearlyChange
                                If ws.Cells(i, 10).Value >= 0 Then
                                ws.Cells(i, 10).Interior.ColorIndex = 4
                                Else: ws.Cells(i, 10).Interior.ColorIndex = 3
                                End If
        Next i

' Max and Min %

        Dim LastRowPercentageChange As Long
          Dim MaxPercent As Double
          Dim MinPercent As Double
          
  
       LastRowPercentageChange = ws.Cells(Rows.Count, 11).End(xlUp).Row
       
      MaxPercent = 0
      MinPercent = 0
      
      For i = 2 To LastRowPercentageChange
                 If MaxPercent < ws.Cells(i, 11).Value Then
                        MaxPercent = ws.Cells(i, 11).Value
                        ws.Cells(2, 17).Value = MaxPercent
                         ws.Cells(2, 17).Value = Format(ws.Cells(2, 17).Value, "Percent")
                          ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                          
                 ElseIf MinPercent > ws.Cells(i, 11).Value Then
                         MinPercent = ws.Cells(i, 11).Value
                        ws.Cells(3, 17).Value = MinPercent
                        ws.Cells(3, 17).Value = Format(ws.Cells(3, 17).Value, "Percent")
                        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                End If
             
     Next i
           
           'Greatest Total Volume
           
           Dim LastRowTotalVol   As Long
           Dim TotalVolumeMax As LongLong
           LastRowTotalVol = ws.Cells(Rows.Count, 12).End(xlUp).Row
        
        TotalVolumeMax = 0
         For i = 2 To LastRowTotalVol
                 If TotalVolumeMax < ws.Cells(i, 12).Value Then
                       TotalVolumeMax = ws.Cells(i, 12).Value
                       ws.Cells(4, 17).Value = TotalVolumeMax
                       ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                       
                       
                       
                      End If
                 
           Next i
                
Next ws


End Sub
