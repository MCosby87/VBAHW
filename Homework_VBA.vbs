Attribute VB_Name = "Module1"
Sub Ticker()
    
    Dim ws As Worksheet
    
    Dim Ticker As String
    
    Dim TickerVolume As Double
    
    Dim Summary_Table_Row As Integer

    For Each ws In Worksheets
    ws.Activate

    Ticker_Volume = 0
    Summary_Table_Row = 2
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Ticker_Volume"
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To LastRow

            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                Ticker = Cells(i, 1).Value
                Ticker_Volume = Ticker_Volume + Cells(i, 7).Value
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ws.Range("J" & Summary_Table_Row).Value = Ticker_Volume
                Summary_Table_Row = Summary_Table_Row + 1
                Total_Stock_Volume = 0

            Else
                TickerVolume = TickerVolume + Cells(i, 7).Value

            End If
              
        Next i

    Next ws

End Sub
