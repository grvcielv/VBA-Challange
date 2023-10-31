# VBA-Challange
For this challenge, I was able to do it with the help of my tutor - Mathew Werth. The part he helped me understand was the following code.
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
    There was another area I received help on to understand yet develop by myself through  ASKBCS is this following code.
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
