VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CalculateYearlyChange()
Dim ws As Worksheet
Dim ticket As String
Dim dia As Date
Dim closingprice As Double
Dim openingprice As Double
Dim yearchange As Double
Dim totalstockvalue As Double
Dim summarytablerow As Integer
Dim Col As Range
Dim maxVal As Double
Dim minVal As Double
Dim maxVal2 As Double


      
For Each ws In Worksheets
With ActiveSheet
openingprice = ws.Cells(2, 3).Value
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
totalstockvalue = 0
ws.Columns(12).ClearContents
ws.Columns(13).ClearContents
ws.Columns(11).ClearContents

j = 2
   For i = 2 To lastrow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticket = ws.Cells(i, 1).Value
        ws.Cells(j, 10).Value = ticket
        closingprice = ws.Cells(i, 6).Value
        yearlychange = closingprice - openingprice
        ws.Cells(j, 11).Value = yearlychange
         If openingprice <> 0 Then
           percentchange = (yearlychange / openingprice)
           ws.Cells(j, 12).Value = percentchange
         Else
           ws.Cells(j, 12).Value = "N/A"
         End If
        openingprice = ws.Cells(i + 1, 3).Value
        j = j + 1
        End If
    Next i
    
    g = 2
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            totalstockvalue = totalstockvalue + ws.Cells(i, 7).Value
            ws.Cells(g, 13).Value = totalstockvalue
            g = g + 1
            totalstockvalue = 0
        Else
            totalstockvalue = totalstockvalue + ws.Cells(i, 7).Value
    End If
    Next i
    
    k = 2
lastrow2 = ws.Cells(Rows.Count, 11).End(xlUp).Row
    For k = 2 To lastrow2
        If IsEmpty(ws.Cells(k, 11).Value) = False Then
            If (ws.Cells(k, 11).Value >= 0) Then
              ws.Cells(k, 11).Interior.ColorIndex = 4 ' Green
            ElseIf (ws.Cells(k, 11).Value < 0) Then
              ws.Cells(k, 11).Interior.ColorIndex = 3 ' Red
            End If
        End If
    Next k
    ws.Range("L2:L" & lastrow2).NumberFormat = "0.00%"
ws.Cells(1, 11).Value = "Yearly Change"
ws.Cells(1, 12).Value = "Percent Change"
ws.Cells(1, 13).Value = "Total Stock Volume"
ws.Cells(1, 10).Value = "Ticker"
    
    For m = 10 To 13
        ws.Cells(1, m).EntireColumn.AutoFit
    Next m

Set Col = ws.Range("K2:K" & lastrow2)
maxVal = Col.Cells(1).Value
minVal = Col.Cells(1).Value
maxVal2 = Col.Cells(1).Value
    For Each cell In Col
        For r = 2 To lastrow2
            If (ws.Cells(r, 12).Value > maxVal) Then
                maxVal = ws.Cells(r, 12).Value
            ElseIf (ws.Cells(r, 12).Value < minVal) Then
                minVal = ws.Cells(r, 12).Value
            End If
            If (ws.Cells(r, 13).Value > maxVal2) Then
                maxVal2 = ws.Cells(r, 13).Value
            End If
        Next r
Next cell
ws.Cells(2, 17).Value = maxVal
ws.Cells(3, 17).Value = minVal
ws.Cells(4, 17).Value = maxVal2
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"
For d = 2 To lastrow2
        If (ws.Cells(2, 17).Value = ws.Cells(d, 12).Value) Then
            ticketforhighpercent = ws.Cells(d, 10).Value
        End If
        If (ws.Cells(3, 17).Value = ws.Cells(d, 12).Value) Then
            ticketforlowpercent = ws.Cells(d, 10).Value
        End If
        If (ws.Cells(4, 17).Value = ws.Cells(d, 13).Value) Then
            ticketforvol = ws.Cells(d, 10).Value
        End If
    Next d
    
ws.Cells(2, 16).Value = ticketforhighpercent
ws.Cells(3, 16).Value = ticketforlowpercent
ws.Cells(4, 16).Value = ticketforvol

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

    For w = 2 To 4
        Cells(w, 15).EntireColumn.AutoFit
    Next w

ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"


 End With
Next ws
End Sub

