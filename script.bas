Dim Ticker As String
Sub zob()

Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double
Dim TotalVolume As Double


'Ticker name as column I
    Dim j As Long
    j = 2
    
    [I1] = "Ticker"
    [J1] = "Yearly Change"
    [K1] = "Percentage Change"
    [L1] = "Total Stock Volume"

'Loop thru all ticker symbols
 
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
    
    'Process first line for a ticker
    If (Ticker <> Cells(i, 1).Value) Then
        If (i <> 2) Then
            Change = ClosePrice - OpenPrice
            Cells(j, 9).Value = Ticker
            Cells(j, 10).Value = Change
            If Change > 0 Then
                    Cells(j, 10).Interior.ColorIndex = 4
                Else
                    Cells(j, 10).Interior.ColorIndex = 3
                    
            End If
            Cells(j, 11).Value = (ClosePrice - OpenPrice) / OpenPrice
            Cells(j, 12).Value = TotalVolume
            j = j + 1
        End If
        Ticker = Cells(i, 1).Value
        OpenPrice = Cells(i, 3).Value
        TotalVolume = Cells(i, 7).Value
    Else
        TotalVolume = TotalVolume + Cells(i, 7).Value
        ClosePrice = Cells(i, 6).Value
    End If
    


Next i



End Sub
