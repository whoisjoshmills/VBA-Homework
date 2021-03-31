Attribute VB_Name = "Module1"
Sub vba_testing()

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Dim ticker As String
Dim row As Integer
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_value As Double
total_stock_value = 0

row = 2
lastrow = Cells(Rows.Count, 1).End(xlUp).row
For i = 2 To lastrow

If Cells(i + 1, 1).Value <> Cells(i, 1) Then
ticker = Cells(i, 1).Value
yearly_change = Cells(i - 261, 3).Value - Cells(i, 6).Value
percent_change = yearly_change / Cells(i - 261, 3).Value
total_stock_value = total_stock_value + Cells(i, 7).Value
Range("I" & row).Value = ticker
Range("J" & row).Value = yearly_change



Range("K" & row).Value = percent_change
Range("K" & row).NumberFormat = "0.00%"
Range("L" & row).Value = total_stock_value


row = row + 1
total_stock_value = 0
'yearly_change =
Else
yearly_change = Cells(i, 3).Value - Cells(i, 6)
total_stock_value = total_stock_value + Cells(i, 7).Value

End If






Next i
End Sub

