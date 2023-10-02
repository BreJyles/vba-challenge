Attribute VB_Name = "Module1"
Sub numbers()
Dim lastrow As Double
Dim ticker As String
Dim yearly_change As Double
Dim percentage_change As Double
Dim total_stock_volume As Double
Dim ticket As String
Dim open_price As Double
Dim close_price As Double
Dim a As Double
Dim openindex As Long
Dim cellnumber As Integer
Dim color_options(2) As Integer
Dim maxincreaseticker As String
Dim maxdecrease As Double
Dim maxincrease As Double
Dim maxdecreaseticker As String
Dim maxvolume As Double
Dim maxvolumeticker As String
Dim ws As Worksheet

'For Each ws In this number Worksheet




lastrow = Cells(Rows.Count, 1).End(xlUp).Row
'header

 Cells(1, 9).Value = "ticker"
 Cells(1, 10).Value = "yearly change"
 Cells(1, 11).Value = "percent change"
 Cells(1, 12).Value = "total stock volume"
 'declare variable
 
 lastrow = Cells(Rows.Count, 1).End(x1up).Row
  num = 2
  stockvalue = 0
  cellnumber = 2
  openindex = 2
  maxincrease = 0
  maxdecrease = 0
  maxvolume = 0
  
  
 
 
For i = 2 To lastrow

stockvalue = stockvalue + Cells(i, 7).Value

'if statments that finds value in the columns
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
 ticker = Cells(1, 1).Value
 openprice = Cells(openindex, 3).Value
 closeprice = Cells(1, 6).Value
 yearly_change = closeprice = openprice
 percent_change = yearly_change / openprice


Cells(1, 9).Value = "ticker"

Cells(1, 10).Value = "yearly_change"

Cells(1, 11).Value = "percent change"

Cells(1, 12).Value = "total stock volume"

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then


ticker = Cells(i, 1).Value


'formatting percent change as percentage
Cells(1, 11).NumberFormat = "0.00%"


Else



total_stock_volume = total_stock_volume + Cells(i, 7).Value

End If





yearly_change = ws.Cells(Rows.Count, 10).End(x1up).Row



For r = 2 To yearly_change

If ws.Cells(r, 10).Value < 0 Then

ws.Cells(r, 10).Interior.ColorIndex = 10

End If


Next r






Range("j" & ticker_row).Value = yearly_change


If open_value = 0 Then

percent_change = yearly_change / open_value
End If

ws.Range("k" & ticker_row).Value = percent_change


ticker_row = ticker_row + 1

total_stock_volume = 0

open_value = ws.Cells(1 + i, 3)


Else


total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value

End If


Next i




yearly_change = ws.Cells(Rows.Count, 10).End(x1up).Row

For r = 2 To yearly_change

If ws.Cells(r, 10).Value < 0 Then

ws.Cells(r, 10).Interior.ColorIndex = 10


End If


Next r




For k = 2 To yearly_change

ws.Range("k2:k" & yearly_change).NumberFormat = "0.00%"

Next k






ws.Cells(2, 15).Value = "max % increase"
ws.Cells(3, 15).Value = "max % decrease"
ws.Cells(4, 15).Value = "max total volume"
ws.Cells(1, 16).Value = "ticker"
ws.Cells(1, 17).Value = "value"


percent_change = ws.Cells(Rows.Count, 11).End(x1up).Row

For s = 2 To percent_change



If ws.Cells(s, 11).Value = Application.WorksheetFunction.Max(ws.Range("k2:k" & percent_change)) Then

ws.Cells(2, 17).Value = ws.Cells(s, 11).Value
ws.Cells(2, 16).Value = ws.Cells(s, 9).Value
ws.Range("Q3").NumberFormat = "0.00%"




ElseIf ws.Cells(s, 11).Value = Application.WorksheetFunction.Min(ws.Range("k2:k" & percent_change)) Then

ws.Cells(3, 17).Value = ws.Cells(s, 11).Value
ws.Cells(3, 16).Value = ws.Cells(s, 9).Value
ws.Range("Q2").NumberFormat = "0.00%"




ElseIf ws.Cells(s, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & percent_change)) Then

ws.Cells(4, 17).Value = ws.Cells(s, 12).Value
ws.Cells(4, 16).Value = ws.Cells(s, 9).Value








End If

Next s


Next ws





End Sub
