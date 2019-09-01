Attribute VB_Name = "Module1"
Sub AnalyzeStocks()
    Dim Ticker, StockDay As String
    Dim DayOpen, DayHigh, DayLow, DayClose, YearOpen, YearClose, YearVolume As Double
    Dim DayVolume As Long
    Dim RowNum As Integer
    
    WriteHeaders
        

End Sub
Sub WriteHeaders()
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
End Sub
