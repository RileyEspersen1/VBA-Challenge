Attribute VB_Name = "Module1"
Sub StockFunction()
    
    For Each ws In Worksheets
    
        Dim WorksheetName As String
        Dim tickerSymbol As String
        Dim yearlyChange As Double
        Dim percentChange As Double
        Dim totalStockVol As Variant
        Dim yearStart As Double
        Dim yearEnd As Double
        Dim printCount As Integer
        Dim count As Integer
        
        
        LastRow = ws.Cells(Rows.count, 1).End(xlUp).Row
    
        'For loop to go through all the sheets ***
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        WorksheetName = ws.Name
        printCounter = 2
        totalStockVol = 0
    
        For i = 2 To LastRow
            
            totalStockVol = totalStockVol + Cells(i, 7).Value
               
            If (ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value) Then
                
                yearStart = ws.Cells(i, 3).Value
            
            End If
                    
            If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                                                                                                                                                                                                                            
                'MsgBox (totalStockVol)
                'yearStart
                'MsgBox (yearStart)
    
                'endYear
                yearEnd = ws.Cells(i, 6).Value
                'MsgBox ("Year end: " + Str(yearEnd))
                                       
                'yearlyChanges
                yearlyChange = yearEnd - yearStart
                'MsgBox ("Yearly Change: " + Str(yearlyChange))
                                      
                'totalStockVol in L Col
                ws.Cells(printCounter, 12).Value = totalStockVol
                
                'prints ticker value in I Col
                ws.Cells(printCounter, 9).Value = ws.Cells(i, 1).Value
                
                'prints yearlyChanges in J col
                ws.Cells(printCounter, 10).Value = Str(yearlyChange)
                
                'changing color based on yearlyChange
                If (yearlyChange >= 0) Then
                    ws.Range("J" & printCounter).Value = yearlyChange
                    ws.Range("J" & printCounter).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & printCounter).Value = yearlyChange
                    ws.Range("J" & printCounter).Interior.ColorIndex = 3
                
                End If
                
                ws.Columns("K").NumberFormat = "##,##0.00%"
                'print percentChange
                If (yearStart > 0) Then
                    
                    percentChange = (yearlyChange / yearStart)
                
                End If
                ws.Cells(printCounter, 11).Value = percentChange
                
                'counters
                printCounter = printCounter + 1
                totalStockVol = 0

            End If

        Next i
        
        'challenges
        
        LastRow2 = ws.Cells(Rows.count, 11).End(xlUp).Row
        
        
        Dim posCurrentVal As Double
        Dim posCurrentTickVal As String
        Dim negCurrentVal As Double
        Dim negCurrentTickVal As String
        
        'biggest positive % change
        posCurrentVal = ws.Cells(2, 11).Value
        posCurrentTickVal = ws.Cells(2, 9).Value
        
        For i = 2 To LastRow2
            If (ws.Cells(i + 1, 11).Value > posCurrentVal) Then
                posCurrentVal = ws.Cells(i, 11).Value
                posCurrentTickVal = ws.Cells(i, 9).Value
            End If
            
        Next i
            
        'biggest negative % change
        negCurrentVal = ws.Cells(2, 11).Value
        negCurrentTickVal = ws.Cells(2, 9).Value
        
        For i = 2 To LastRow2
            If (ws.Cells(i + 1, 11).Value < negCurrentVal) Then
                negCurrentVal = ws.Cells(i, 11).Value
                negCurrentTickVal = ws.Cells(i, 9).Value
            End If
            
        Next i
        
        'greatest volume
        currentVol = ws.Cells(2, 12).Value
        volTickVal = ws.Cells(2, 9).Value

        For i = 2 To LastRow2
            If (ws.Cells(i + 1, 12).Value > currentVol) Then
                currentVol = ws.Cells(i, 12).Value
                volTickVal = ws.Cells(i, 9).Value
            End If
            
        Next i
            
        ws.Cells(1, 16).Value = “Ticker”
        ws.Cells(1, 17).Value = “Value”
        
        '% increased data
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(2, 16).Value = posCurrentTickVal
        ws.Cells(2, 17).Value = posCurrentVal
        
        '% decreased data
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = negCurrentTickVal
        ws.Cells(3, 17).Value = negCurrentVal
        
        'Greatest Volume
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = volTickVal
        ws.Cells(4, 17).Value = currentVol
        
        'formatting
        ws.Cells(2, 17).NumberFormat = "##,##0.00%"
        ws.Cells(3, 17).NumberFormat = "##,##0.00%"

    Next ws
      
End Sub
    
