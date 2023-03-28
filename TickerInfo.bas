Attribute VB_Name = "Module1"
Sub TickerInfo()

' Initializing variables
Dim ws As Worksheet
For Each ws In Worksheets

Dim i, rowOutput, lastRow As Long
Dim priOpen, priClose, priDiff, priDiffPer, stVol As Double
Dim topPer, botPer, topVol As Double
Dim liveTicker, topPerTicker, botPerTicker, topVolTicker As String

    ' Organizing output columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

    ' Assigning values to variables
    rowOutput = 2
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    priOpen = ws.Cells(2, 3).Value
    stVol = ws.Cells(2, 7).Value
    liveTicker = ws.Cells(2, 1).Value
    topPerTicker = ws.Cells(2, 1).Value
    botPerTicker = ws.Cells(2, 1).Value
    topVolTicker = ws.Cells(2, 1).Value
    topPer = 0
    botPer = 0
    topVol = 0
    
    
    
        ' Scanning live worksheet for data
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
              stVol = stVol + ws.Cells(i, 7).Value
            
            ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
              priClose = ws.Cells(i, 6).Value
              priDiff = priClose - priOpen
              
                ' Exception for 0, checking & updating % values
                If priDiff <> 0 Then
                  priDiffPer = (priDiff / priOpen) * 100
                        If (priDiffPer > topPer) Then
                          topPer = priDiffPer
                          topPerTicker = ws.Cells(i, 1).Value
                        ElseIf (priDiffPer < botPer) Then
                          botPer = priDiffPer
                          botPerTicker = ws.Cells(i, 1).Value
                        End If
                Else: ws.Cells(rowOutput, 11).Value = 0
                End If
                
                ' Checking & updating total vol values
                If stVol > topVol Then
                    topVol = stVol
                    topVolTicker = ws.Cells(i, 1).Value
                End If
                
            ' Printing data to output columns, conditional formatting
            ws.Cells(rowOutput, 9).Value = liveTicker
            ws.Cells(rowOutput, 10).Value = priDiff
                If priDiff > 0 Then
                  ws.Cells(rowOutput, 10).Interior.ColorIndex = 4
                ElseIf priDiff < 0 Then
                  ws.Cells(rowOutput, 10).Interior.ColorIndex = 3
                End If
            ws.Cells(rowOutput, 11).Value = Str(priDiffPer) + "%"
            ws.Cells(rowOutput, 12).Value = stVol
            rowOutput = rowOutput + 1
            
            
            ' Reassign values from beginning of new ticker
            stVol = 0
            priOpen = ws.Cells(i + 1, 3).Value
            liveTicker = ws.Cells(i + 1, 1).Value
            
            End If
        Next i
        
        ' Printing top/ bot performers from entire sheet, formatting
        ws.Cells(2, 16).Value = topPerTicker
        ws.Cells(3, 16).Value = botPerTicker
        ws.Cells(4, 16).Value = topVolTicker
        ws.Cells(2, 17).Value = Str(topPer) + "%"
        ws.Cells(3, 17).Value = Str(botPer) + "%"
        ws.Cells(4, 17).Value = topVol
        ws.Columns("A:Z").AutoFit

Next ws
End Sub


