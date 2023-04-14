Attribute VB_Name = "Module1"
Sub tickerBoy()



Dim ws As Worksheet

Dim lastRow As String

Dim summaryTableRow As Integer

Dim percentChange As Double

Dim stockVolume As LongLong

Dim percentChangeMax As Double

Dim percentChangeMaxTicker As String

Dim wsMax As Double

Dim header(9) As String




    For Each ws In ActiveWorkbook.Worksheets
    
    
    
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    percentChangeLastRow = ws.Cells(Rows.Count, "M").End(xlUp).Row
    
    summaryTableRow = 2
    
    stockVolume = 0
    
    percentChangeMax = 0
    
    percentChangeMin = 0
    
    maxTotalVolume = 0
    
    openVal = ws.Cells(2, 3).Value
    
    percentChangeLastRow = ws.Cells(Rows.Count, "M").End(xlUp).Row
    
    
    
    ws.Cells(1, 11) = "Ticker"
    
    ws.Cells(1, 12) = "Yearly Change"
    
    ws.Cells(1, 13) = "Percent Change"
    
    ws.Cells(1, 14) = "Total Stock Volume"
    
    ws.Cells(1, 19) = "Ticker"
    
    ws.Cells(1, 20) = "Value"
    
    ws.Cells(2, 18) = "Greatest % Increase"
    
    ws.Cells(3, 18) = "Greatest % Decrease"
    
    ws.Cells(4, 18) = "Greatest Total Volume"
    
        
        
        For j = 2 To lastRow
        
        
    
            If ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then
            
            
            
                stockVolume = stockVolume + ws.Cells(j, 7).Value

                ticker = ws.Cells(j, 1).Value
            
                closeVal = ws.Cells(j, 6).Value
            
                yearlyChange = closeVal - openVal
                
                
                
                If yearlyChange > 0 Then
                
                    ws.Range("L" & summaryTableRow).Interior.ColorIndex = 4
                    
                Else
                
                    ws.Range("L" & summaryTableRow).Interior.ColorIndex = 3
                    
                End If
                
                
                
                For m = 2 To percentChangeLastRow
                    
                    If ws.Cells(m, 13).Value > percentChangeMax Then
                        
                        percentChangeMax = ws.Cells(m, 13).Value
                        
                        percentChangeMaxTicker = ws.Cells(m, 11).Value
                    
                    End If
                    
                Next m
                
                
                
                For m = 2 To percentChangeLastRow
                    
                    If ws.Cells(m, 13).Value < percentChangeMin Then
                        
                        percentChangeMin = ws.Cells(m, 13).Value
                        
                        percentChangeMinTicker = ws.Cells(m, 11).Value
                        
                    End If
                    
                Next m
                
                
                
                For m = 2 To percentChangeLastRow
                    
                    If ws.Cells(m, 14).Value > maxTotalVolume Then
                        
                        maxTotalVolume = ws.Cells(m, 14).Value
                        
                        totalVolumeTicker = ws.Cells(m, 11).Value
                        
                    End If
                    
                Next m
                
                
                
                percentChange = ((closeVal - openVal) / openVal)
        
                ws.Range("K" & summaryTableRow).Value = ticker
             
                ws.Range("L" & summaryTableRow).Value = yearlyChange
            
                ws.Range("M" & summaryTableRow).Value = percentChange
                
                ws.Range("M" & summaryTableRow).NumberFormat = "0.00%"
            
                ws.Range("N" & summaryTableRow).Value = stockVolume

                stockVolume = 0
            
                openVal = ws.Cells(j + 1, 3).Value
            
                summaryTableRow = summaryTableRow + 1
            
            
            Else

                stockVolume = stockVolume + ws.Cells(j, 7).Value
        
        
            End If
    
        Next j
        
                
                ws.Cells(2, 19).Value = percentChangeMaxTicker
                
                ws.Cells(3, 19).Value = percentChangeMinTicker
                
                ws.Cells(4, 19).Value = totalVolumeTicker
                
                ws.Cells(2, 20).Value = percentChangeMax
                
                ws.Cells(3, 20).Value = percentChangeMin
                
                ws.Cells(3, 20).NumberFormat = "0.00%"
                
                ws.Cells(4, 20).Value = maxTotalVolume
                
                ws.UsedRange.EntireColumn.AutoFit
                
                ws.UsedRange.EntireRow.AutoFit
    
    Next ws
    
End Sub

