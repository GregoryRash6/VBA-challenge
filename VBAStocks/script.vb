Sub Stock_Market_Analysis():

For Each ws In Worksheets

Dim LastRow As Long
Dim Ticker As String
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
Dim Yearly_Open As Double
Dim Yearly_Close As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Previous_Amount As Long
Previous_Amount = 2
Dim Ticker_Volume As Double
Ticker_Volume = 0
Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0
Dim Greatest_Increase As Double
Greatest_Increase = 0
Dim Greatest_Decrease As Double
Greatest_Decrease = 0
Dim Greatest_Total_Volume As Double
Greatest_Total_Volume = 0

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Total Stock Volume"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow

Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

Ticker = ws.Cells(i, 1).Value
ws.Range("I" & Summary_Table_Row).Value = Ticker

Yearly_Open = ws.Range("C" & Previous_Amount)
Yearly_Close = ws.Range("F" & i)
Yearly_Change = Yearly_Close - Yearly_Open
ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

If Yearly_Open = 0 Then
Percent_Change = 0
Else
Yearly_Open = ws.Range("C" & Previous_Amount)
Percent_Change = Yearly_Change / Yearly_Open
End If
ws.Range("K" & Summary_Table_Row).Value = Percent_Change
ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

ws.Range("L" & Summary_Table_Row).Value = Ticker_Volume
Ticker_Volume = 0

If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
ws.Range("J" & Summary_Table_Row).Interior.Color = vbGreen
Else
ws.Range("J" & Summary_Table_Row).Interior.Color = vbRed
End If

Summary_Table_Row = Summary_Table_Row + 1
Previous_Amount = i + 1

End If

Next i 

LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row

For i = 2 To LastRow

If ws.Range("K" & i).Value > ws.Range("P2").Value Then
ws.Range("P2").Value = ws.Range("K" & i).Value
ws.Range("O2").Value = ws.Range("I" & i).Value
End If
ws.Range("P2").NumberFormat = "0.00%"

If ws.Range("K" & i).Value < ws.Range("P3").Value Then
ws.Range("P3").Value = ws.Range("K" & i).Value
ws.Range("O3").Value = ws.Range("I" & i).Value
End If
ws.Range("P3").NumberFormat = "0.00%"

If ws.Range("L" & i).Value > ws.Range("P4").Value Then
ws.Range("P4").Value = ws.Range("L" & i).Value
ws.Range("O4").Value = ws.Range("I" & i).Value
End If

Next i
ws.Columns.AutoFit
Next ws

End Sub