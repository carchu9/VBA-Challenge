Sub VBAChallenge():

Dim Ticker As String

Dim TotalChange As Double
TotalChange = 0

Dim TotalVolume As Double
TotalVolume = 0

Dim SummaryTableRow As Integer
SummaryTableRow = 2

Dim YearChange As Double
YearChange = 0

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

OpenPrice = Cells(2, 3).Value
    
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

    'loop through all rows
    For i = 2 To lastrow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            Ticker = Cells(i, 1).Value
            Range("I" & SummaryTableRow).Value = Ticker
        
            ClosePrice = Cells(i, 6).Value
                
            TotalChange = (ClosePrice - OpenPrice) / OpenPrice
            Range("K" & SummaryTableRow).Value = TotalChange
            
            YearChange = ClosePrice - OpenPrice
            Range("J" & SummaryTableRow).Value = YearChange
            
            OpenPrice = Cells(i + 1, 3).Value
       
            TotalVolume = TotalVolume + Cells(i, 7).Value
            Range("L" & SummaryTableRow).Value = TotalVolume
            
            Range("K" & SummaryTableRow).NumberFormat = "0.00%"
            
                If YearChange < 0 Then
                Range("J" & SummaryTableRow).Interior.ColorIndex = 3
            
                ElseIf YearChange > 0 Then
                Range("J" & SummaryTableRow).Interior.ColorIndex = 4
            
                Else
                Range("J" & SummaryTableRow).Interior.ColorIndex = 0
            
                End If
            
                    If TotalChange < 0 Then
                    Range("K" & SummaryTableRow).Interior.ColorIndex = 3
            
                    ElseIf TotalChange > 0 Then
                    Range("K" & SummaryTableRow).Interior.ColorIndex = 4
            
                    Else
                    Range("K" & SummaryTableRow).Interior.ColorIndex = 0
                    End If
        
            
            SummaryTableRow = SummaryTableRow + 1
            
            TotalVolume = 0
        Else
            TotalVolume = TotalVolume + Cells(i, 7).Value
            
        End If
            
    Next i

End Sub