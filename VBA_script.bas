Attribute VB_Name = "Module1"
Sub stockcopy()

'Dim all variables

Dim i As Long
Dim DistinctTicker As String
Dim OpeningVal As Double
Dim Volume As Double
Dim TickerCount As Long
Dim ClosingVal As Double
Dim NewRowCount As Long
Dim ws As Worksheet
Dim TotalRecords As Long
Dim Random As Long
Dim MaxUp As Double
Dim MaxDown As Double
Dim MaxVol As Double
Dim Change As Double



For Each ws In Sheets
'For Ticker Output
'Label columns

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Volume"
ws.Cells(1, 15).Value = "Greatest % Increase"
ws.Cells(1, 16).Value = "Greatest % Decrease"
ws.Cells(1, 17).Value = "Greatest Total Volume"




TickerCount = 0
NewRowCount = 1
Volume = 0
TotalRecords = ws.Cells(Rows.Count, 1).End(xlUp).Row
MaxUp = 0
MaxDown = 0
MaxVol = 0

For i = 1 To TotalRecords
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    
    'abstracting values
    
        DistinctTicker = ws.Cells(i + 1, 1).Value
        OpeningVal = ws.Cells(i + 1, 3).Value
        Random = WorksheetFunction.CountIf(ws.Range("A2:A" & TotalRecords), DistinctTicker)
        Volume = Volume + ws.Cells(i + 1, 7).Value
        TickerCount = TickerCount + WorksheetFunction.CountIf(ws.Range("A2:A" & TotalRecords), DistinctTicker)
        ClosingVal = ws.Cells(TickerCount + 1, 6).Value
        NewRowCount = NewRowCount + 1

        
        
        'assigning values
        
        ws.Cells(NewRowCount, 12).Value = Volume
        ws.Cells(NewRowCount, 9).Value = DistinctTicker
        ws.Cells(NewRowCount, 10).Value = ClosingVal - OpeningVal
        
        ' if to check for div by zero
        If OpeningVal <> 0 Then
            ws.Cells(NewRowCount, 11).Value = (ClosingVal / OpeningVal) - 1
        
        Else: ws.Cells(NewRowCount, 11).Value = 0
        
        End If
        
        'Color Coding
        
        
       If (ws.Cells(NewRowCount, 10).Value >= 0) Then
            ws.Cells(NewRowCount, 10).Interior.Color = vbGreen
       
       Else: ws.Cells(NewRowCount, 10).Interior.Color = vbRed
       
       End If
       
        Change = ws.Cells(NewRowCount, 11).Value
        
        'max up and down
        
        If Change >= 0 And Change >= MaxUp Then
            MaxUp = Change
        ElseIf Change < 0 And Change < MaxDown Then
            MaxDown = Change
        End If
        
        'max volume
        
        If Volume >= MaxVol Then
            MaxVol = Volume
        End If
        
       
       Volume = 0
        
        Else:  Volume = Volume + ws.Cells(i + 1, 7).Value
     
    End If
    
Next i

ws.Cells(2, 15).Value = MaxUp
ws.Cells(2, 16).Value = MaxDown
ws.Cells(2, 17).Value = MaxVol

Next ws

End Sub



