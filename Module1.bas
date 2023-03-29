Attribute VB_Name = "Module1"
Option Explicit

Sub main()
Dim i As Long
Dim openprice, closingprice, totalVolume As Double
Dim tickerCnt As Long
Dim ws As Worksheet
Dim topIncreaseTicker, lowIncreaseTicker, topVolumeTicker As String
Dim topIncreaseValue, lowIncreaseValue, topVolumeValue As Double

Call clearWorksheets

'Initialize
tickerCnt = 0
totalVolume = 0

For Each ws In Worksheets
    Sheets(ws.Name).Activate
    Cells(1, "A").Select
    
    'Create Summary Header
    With Cells(1, 9)
        .value = "Ticker"
        .Offset(, 1).value = "Yearly Change"
        .Offset(, 2).value = "Percent Change"
        .Offset(, 3).value = "Total Stock Volume"
    End With

    Range("I1", "L1").Columns.AutoFit

    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row

        tickerCnt = tickerCnt + 1
        totalVolume = totalVolume + Cells(i, "A").Offset(, 6).value

        If tickerCnt = 1 Then
            openprice = Cells(i, "A").Offset(, 2).value
        End If
    
        'Create Summary Row
        If Cells(i, "A").value <> Cells(i, "A").Offset(1).value Then
            closingprice = Cells(i, "A").Offset(, 5).value
        
            With Cells(Cells(Rows.Count, 9).End(xlUp).Row + 1, 9)
                .value = Cells(i, "A").value
                .Offset(, 1).value = closingprice - openprice
    
                If closingprice - openprice < 0 Then
                    .Offset(, 1).Interior.Color = RGB(255, 0, 0)
                    '.Offset(, 2).Interior.Color = RGB(255, 0, 0)
                Else
                    .Offset(, 1).Interior.Color = RGB(0, 255, 0)
                    '.Offset(, 2).Interior.Color = RGB(0, 255, 0)
                End If
    
                .Offset(, 2).value = FormatPercent((closingprice - openprice) / openprice, 2)
                .Offset(, 3).value = totalVolume

            End With
            
            tickerCnt = 0
            totalVolume = 0
        
        End If
    
    Next i
    
    'Create Statistics Headers
    Cells(2, "O").value = "Greatest % Increase"
    Cells(3, "O").value = "Greatest % Decrease" '
    Cells(4, "O").value = "Greatest Total Volume"
    Cells(1, "P").value = "Ticker"
    Cells(1, "Q").value = "Value"
       
    'Initialize using first row of Summary Data
    topIncreaseValue = Cells(2, "K").value
    lowIncreaseValue = Cells(2, "K").value
    topVolumeValue = Cells(2, "L").value
    
    'Create Statistics Rows
    For i = 3 To Cells(Rows.Count, 9).End(xlUp).Row
        If Cells(i, "K").value > topIncreaseValue Then
            topIncreaseTicker = Cells(i, "I").value
            topIncreaseValue = Cells(i, "K").value
        ElseIf Cells(i, "K").value = topIncreaseValue Then
            topIncreaseTicker = topIncreaseTicker & ", " & Cells(i, "I").value
        End If
        
        If Cells(i, "K").value < lowIncreaseValue Then
            lowIncreaseTicker = Cells(i, "I").value
            lowIncreaseValue = Cells(i, "K").value
        ElseIf Cells(i, "K").value = lowIncreaseValue Then
            lowIncreaseTicker = lowIncreaseTicker & ", " & Cells(i, "I").value
        End If
        
        If Cells(i, "L").value > topVolumeValue Then
            topVolumeTicker = Cells(i, "I").value
            topVolumeValue = Cells(i, "L").value
        ElseIf Cells(i, "L").value = topVolumeValue Then
            topVolumeTicker = topVolumeTicker & ", " & Cells(i, "I").value
        End If
        
        Cells(2, "P").value = topIncreaseTicker
        Cells(2, "Q").value = FormatPercent(topIncreaseValue, 2)
        
        Cells(3, "P").value = lowIncreaseTicker
        Cells(3, "Q").value = FormatPercent(lowIncreaseValue, 2)
        
        Cells(4, "P").value = topVolumeTicker
        Cells(4, "Q").value = topVolumeValue
        
    Next i
    
    'Perform some housekeeping
    Cells(1, 1).Select
    
Next ws

Call goHome

End Sub

Sub clearWorksheets()
Dim ws As Worksheet
Dim lastRow As Long

For Each ws In Worksheets
    Sheets(ws.Name).Activate
    lastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    Range("I1:Q" & lastRow).Delete
Next ws

Call goHome

End Sub

Sub goHome()
    
    Application.Goto Sheets(1).Cells(1, 1)

End Sub
