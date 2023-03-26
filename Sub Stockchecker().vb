Sub Stockchecker()

For Each ws In ActiveWorkbook.Worksheets

'sorts data for consistency
ws.Range("A:G").Sort Key1:=ws.Range("B1"), Header:=xlYes
ws.Range("A:G").Sort Key1:=ws.Range("A1"), Header:=xlYes

'sets headers and primes data
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest%Increase"
ws.Range("Q2").Value = 0
ws.Range("O3").Value = "Greatest%Decrease"
ws.Range("Q3").Value = 0
ws.Range("O4").Value = "Greatest Total Volume"
ws.Cells(2, 9).Value = ws.Cells(2, 1).Value
ws.Cells(2, 10).Value = ws.Cells(2, 3).Value

Dim lrow As Long
Dim trow As Integer
Dim ttracker As Integer
Dim openp As Double

trow = 2
lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'MsgBox (lrow)...used for testing

'counts total ticker instances
ttracker = 0
openp = ws.Cells(2, 3).Value

For i = 2 To (lrow + 1)
    
    'Dim closep As Double
    
    
    

    'MsgBox (ttracker)....uesed for testing
    

    If ws.Cells(i, 1).Value <> ws.Cells(trow, 9).Value Then
    trow = trow + 1
    
    'closep = CDbl(ws.cells(i - 1, 6).Value)
    'openp = CDbl(ws.cells(i - (ttracker - 1), 3).Value)
    
    'ticker symbol add
    ws.Cells(trow, 9).Value = ws.Cells(i, 1).Value
    'open add
    ws.Cells(trow, 10).Value = ws.Cells(i, 3).Value
    
    'set volume
    ws.Cells(trow, 12).Value = ws.Cells(i, 7).Value
    'yearly change difference and change cell color
    ws.Cells(trow - 1, 10).Value = ws.Cells(i - 1, 6).Value - ws.Cells(trow - 1, 10).Value
        If ws.Cells(trow - 1, 10).Value >= 0 Then
        ws.Cells(trow - 1, 10).Interior.ColorIndex = 4
        Else
        ws.Cells(trow - 1, 10).Interior.ColorIndex = 3
        End If
    'update percentage, greatest, least, and total volume tracker
    'MsgBox (openp)
    ws.Cells(trow - 1, 11).Value = (ws.Cells(i - 1, 6).Value - openp) / openp
        If ws.Cells(trow - 1, 11).Value >= ws.Range("Q2").Value Then
        ws.Range("Q2").Value = ws.Cells(trow - 1, 11).Value
        ws.Range("P2").Value = ws.Cells(trow - 1, 9).Value
        End If
        If ws.Cells(trow - 1, 11).Value <= ws.Range("Q3").Value Then
        ws.Range("Q3").Value = ws.Cells(trow - 1, 11).Value
        ws.Range("P3").Value = ws.Cells(trow - 1, 9).Value
        End If
        If ws.Cells(trow - 1, 12).Value > ws.Range("Q4").Value Then
        ws.Range("Q4").Value = ws.Cells(trow - 1, 12).Value
        ws.Range("P4").Value = ws.Cells(trow - 1, 9).Value
        End If
        openp = ws.Cells(i, 3).Value
    
        ttracker = ttracker - ttracker
    'add volume to total
    Else: ws.Cells(trow, 12).Value = ws.Cells(trow, 12).Value + ws.Cells(i, 7).Value
    ttracker = ttracker + 1
    
    End If
   
Next i

'formatting ws.column for better readability
ws.Columns("I:I").EntireColumn.AutoFit
ws.Columns("J:J").EntireColumn.AutoFit
ws.Columns("K:K").EntireColumn.AutoFit
ws.Columns("K:K").NumberFormat = "0.00%"
ws.Columns("L:L").EntireColumn.AutoFit
ws.Columns("O:O").EntireColumn.AutoFit
ws.Columns("P:P").EntireColumn.AutoFit
ws.Columns("Q:Q").EntireColumn.AutoFit
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").NumberFormat = "0.00%"

Next ws

End Sub