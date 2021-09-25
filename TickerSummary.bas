Attribute VB_Name = "Module1"
Sub TIckerSummary()

    Dim TickerSymbol As String
    
    Dim TotalVolume As Double
    
    Dim OpeningPrice As Double
    
    Dim ClosingPrice As Double
    
    Dim SummaryRow As Double
    
    SummaryRow = 2
    
    OpeningPrice = Range("C2").Value
    
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        TickerSymbol = Cells(i, 1).Value
        
        TotalVolume = TotalVolume + Cells(i, 7).Value
        
        ClosingPrice = Cells(i, 6).Value
        
        Cells(SummaryRow, 9).Value = TickerSymbol
        
        Cells(SummaryRow, 12).Value = TotalVolume
        
        Cells(SummaryRow, 10).Value = ClosingPrice - OpeningPrice
        
        If Cells(SummaryRow, 10).Value < 0 Then
        
            Cells(SummaryRow, 10).Interior.ColorIndex = 3
            
        Else
            Cells(SummaryRow, 10).Interior.ColorIndex = 4
            
        End If
        
        If OpeningPrice <> 0 Then
        
            Cells(SummaryRow, 11).Value = (ClosingPrice - OpeningPrice) / OpeningPrice
        Else
            Cells(SummaryRow, 11).Value = 0
        End If
        
        SummaryRow = SummaryRow + 1
        
        TotalVolume = 0
        
        OpeningPrice = Cells(i + 1, 3).Value
        
        Else
        
        TotalVolume = TotalVolume + Cells(i, 7).Value
        
        End If
        
        Next i
        
        
        
End Sub
