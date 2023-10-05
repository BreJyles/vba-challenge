VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub numbers()
Dim lastrow As Double
Dim ticker As String
Dim yearly_change As Double
Dim percentage_change As Double
Dim total_stock_volume As Double
Dim ticket As String
Dim open_price As Double
Dim close_price As Double
Dim m As Double
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
Dim num As Double
Dim total As Double



For Each ws In Worksheets


ws.Range("i1") = "ticker"
ws.Range("j1") = "yearly_change"
ws.Range("k1") = "percent_change"
ws.Range("l1") = "total_stock_volume"




lastrow = Cells(Rows.Count, 1).End(xlUp).Row


'header

ws.Range("p1") = "ticker"
ws.Range("q1") = "value"
ws.Range("o2") = "greatest % increase"
ws.Range("o3") = "greatest % decrease"
ws.Range("o4") = "greatest total volume"

 
 
 
 
 
 'declare variable
 
 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  num = 2
  stockvalue = 0
  cellnumber = 2
  openindex = 2
  maxincrease = 0
  maxdecrease = 0
  maxvolume = 0
  open_price = ws.Range("c2")
  
  
 
 
For i = 2 To lastrow

total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value



'if statments that finds value in the columns

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
ticker = ws.Cells(i, 1)
close_price = ws.Cells(i, 6)
yearly_change = close_price - open_price
percentage_change = yearly_change / open_price

ws.Cells(num, 9) = ticker
ws.Cells(num, 10) = yearly_change
ws.Cells(num, 11) = percentage_change
ws.Cells(num, 11).NumberFormat = "0.00%"
ws.Cells(num, 12) = total_stock_volume



'reset variable for next loop

open_price = ws.Cells(i + 1, 3)
num = num + 1
total_stock_volume = 0
End If



Next i



'conditional formatting

For j = 2 To lastrow
Dim ranges As Range
Set ranges = ws.Range("j" & j)

If ranges > 0 Then
ranges.Interior.ColorIndex = 4

ElseIf ranges < 0 Then
ranges.Interior.ColorIndex = 3

End If

Next j


'defining ranges
Change = ws.Range("k:k")
totalvolume = ws.Range("l:l")

'setting variables

maxincrease = WorksheetFunction.Max(Change)
maxdecrease = WorksheetFunction.Min(Change)
maxvolume = WorksheetFunction.Max(totalvolume)


For m = 2 To lastrow

If ws.Cells(m, 11).Value = maxincrease Then
ws.Range("p2") = ws.Cells(m, 9)
ws.Range("q2") = ws.Cells(m, 11)
ws.Range("q2").NumberFormat = "0.00%"
ElseIf ws.Cells(m, 11).Value = maxdecrease Then

ws.Range("p3") = ws.Cells(m, 9)
ws.Range("q3") = ws.Cells(m, 11)
ws.Range("q3").NumberFormat = "0.00%"

ElseIf ws.Cells(m, 12).Value = maxvolume Then
ws.Range("p4") = ws.Cells(m, 9)
ws.Range("q4") = ws.Cells(m, 12)
End If

Next m

Next ws


End Sub

