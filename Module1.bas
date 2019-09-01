Attribute VB_Name = "Module1"
Option Explicit
'Define columns
Public Const OriginalTickerCol = 1
Public Const DateCol = 2
Public Const OpenCol = 3
Public Const HighCol = 4
Public Const LowCol = 5
Public Const CloseCol = 6
Public Const VolumeCol = 7
Public Const TickersCol As Integer = 9
Public Const YearOpenCol As Integer = 10
Public Const YearCloseCol As Integer = 11
Public Const YearlyChangeCol As Integer = 12
Public Const PercentChangeCol As Integer = 13
Public Const TotalVolumeCol As Integer = 14
Public Const GreatestLabelCol As Integer = 16
Public Const Greatest2016TickerCol As Integer = 17
Public Const Greatest2016ValueCol As Integer = 18
Public Const Greatest2015TickerCol As Integer = 19
Public Const Greatest2015ValueCol As Integer = 20
Public Const Greatest2014TickerCol As Integer = 21
Public Const Greatest2014ValueCol As Integer = 22
Public Const GreatestOverallTickerCol As Integer = 23
Public Const GreatestOverallValueCol As Integer = 24
Public Const HeaderRow As Integer = 1
Public Const GreatestIncreaseRow As Integer = 2
Public Const GreatestDecreaseRow As Integer = 3
Public Const GreatestVolumeRow As Integer = 4
Public Const CurrentTickerRow As Integer = 6

Sub AnalyzeStocks()

    Dim ws As Worksheet
    Dim SheetNum, Stocksheets, Tickers, OriginalRowNum, TickerRow, GreatestTickerCol, GreatestValueCol As Integer
    Dim SheetName, StockTicker, PreviousTicker, NextTicker As String
    Dim FirstRow, LastRow, FirstOpenFound As Boolean
    Dim RowOpen, YearOpen, YearClose, YearChange, YearPercentChange, GreatestIncrease, GreatestDecrease, OverallGreatestIncrease, OverallGreatestDecrease As Double
    Dim Volume, TotalVolume, GreatestVolume, OverallGreatestVolume As LongLong
    
    'Initialize sheet variables
    OverallGreatestIncrease = 0
    OverallGreatestDecrease = 0
    OverallGreatestVolume = 0
    TickerRow = 1
      
    'Loop through Worksheets
    Stocksheets = ActiveWorkbook.Worksheets.Count
    For SheetNum = 1 To Stocksheets
        Set ws = ActiveWorkbook.Worksheets(SheetNum)
        With ws
                               
            'Write ticker headers
            .Range("H:Z").Value = ""
            .Cells(HeaderRow, TickersCol).Value = "Ticker"
            .Cells(HeaderRow, YearlyChangeCol).Value = "Yearly Change"
            .Cells(HeaderRow, PercentChangeCol).Value = "Percent Change"
            .Cells(HeaderRow, TotalVolumeCol).Value = "Total Stock Volume"
            .Cells(HeaderRow, YearOpenCol).Value = "Year Open"
            .Cells(HeaderRow, YearCloseCol).Value = "Year Close"
            
            'Write greatest change headers
            SheetName = .Name
            Select Case Trim(SheetName)
                Case "2016"
                    GreatestTickerCol = Greatest2016TickerCol
                    GreatestValueCol = Greatest2016ValueCol
                Case "2015"
                    GreatestTickerCol = Greatest2015TickerCol
                    GreatestValueCol = Greatest2015ValueCol
                Case "2014"
                    GreatestTickerCol = Greatest2014TickerCol
                    GreatestValueCol = Greatest2014ValueCol
            End Select 'SheetNum
            ActiveWorkbook.Worksheets(1).Cells(HeaderRow, GreatestTickerCol).Value = SheetName + " Ticker"
            ActiveWorkbook.Worksheets(1).Cells(HeaderRow, GreatestValueCol).Value = SheetName + " Value"
            If SheetNum = 1 Then
                .Cells(HeaderRow + 1, GreatestLabelCol).Value = "Greatest % Increase"
                .Cells(HeaderRow + 2, GreatestLabelCol).Value = "Greatest % Decrease"
                .Cells(HeaderRow + 3, GreatestLabelCol).Value = "Greatest Total Volume"
                .Cells(HeaderRow, GreatestOverallTickerCol).Value = "All Time Ticker"
                .Cells(HeaderRow, GreatestOverallValueCol).Value = "All Time Value"
            End If 'SheetNum = 1
            
            'Initialize annual greatest change variables
            GreatestIncrease = 0
            GreatestDecrease = 0
            GreatestVolume = 0
                      
            'Initiailize ticker day loop variables
            OriginalRowNum = 2
            TickerRow = 2
            Tickers = 0
            TotalVolume = 0
            FirstOpenFound = False
            YearOpen = 0
            YearClose = 0
            StockTicker = .Cells(OriginalRowNum, OriginalTickerCol).Value
            
            'Continue until all transactions have been read
            Do While Len(StockTicker) > 0
                StockTicker = .Cells(OriginalRowNum, OriginalTickerCol).Value
                PreviousTicker = .Cells(OriginalRowNum - 1, OriginalTickerCol).Value
                NextTicker = .Cells(OriginalRowNum + 1, OriginalTickerCol).Value
                'First row has a header or different ticker before it
                FirstRow = (StockTicker <> PreviousTicker)
                'Last row has a blank row or different ticker after it
                LastRow = (StockTicker <> NextTicker)
                'Aggregate transaction volume
                Volume = .Cells(OriginalRowNum, VolumeCol).Value
                TotalVolume = TotalVolume + Volume
                'Check for Year Open if not found, yet, and update if found
                If Not FirstOpenFound Then
                    RowOpen = .Cells(OriginalRowNum, OpenCol).Value
                    FirstOpenFound = (RowOpen > 0)
                    If FirstOpenFound Then
                        YearOpen = RowOpen
                        .Cells(TickerRow, YearOpenCol).Value = YearOpen
                    End If 'FirstOpenFound
                End If 'Not FirstOpenFound
                If FirstRow Then
                    'Update ticker debug display
                    ActiveWorkbook.Worksheets(1).Cells(CurrentTickerRow, GreatestTickerCol).Value = StockTicker
                    'Increment ticker counts and add ticker to the distinct Tickers list
                    Tickers = Tickers + 1
                    .Cells(TickerRow, TickersCol).Value = StockTicker
                Else 'Not First Row
                    If LastRow Then
                        'Year close is in last row, also calculate annual change and % change, and finish iterating current ticker.
                        YearClose = .Cells(OriginalRowNum, CloseCol).Value
                        .Cells(TickerRow, YearCloseCol).Value = YearClose
                        .Cells(TickerRow, TotalVolumeCol).Value = TotalVolume
                        YearChange = YearClose - YearOpen
                        .Cells(TickerRow, YearlyChangeCol).Value = YearChange
                        If YearOpen > 0 Then 'calculate and analyze % change, otherwise % change undefined
                            'Calculate % change
                            YearPercentChange = YearChange / YearOpen
                            .Cells(TickerRow, PercentChangeCol).Value = YearPercentChange
                            'Analyze % Change
                            If YearPercentChange > GreatestIncrease Then
                                ActiveWorkbook.Worksheets(1).Cells(GreatestIncreaseRow, GreatestTickerCol).Value = StockTicker
                                ActiveWorkbook.Worksheets(1).Cells(GreatestIncreaseRow, GreatestValueCol).Value = YearPercentChange
                                GreatestIncrease = YearPercentChange
                                If YearPercentChange > OverallGreatestIncrease Then
                                    ActiveWorkbook.Worksheets(1).Cells(GreatestIncreaseRow, GreatestOverallTickerCol).Value = StockTicker
                                    ActiveWorkbook.Worksheets(1).Cells(GreatestIncreaseRow, GreatestOverallValueCol).Value = YearPercentChange
                                    OverallGreatestIncrease = YearPercentChange
                                End If 'Overall Greatest % change
                            Else 'Not greatest increase
                                If YearPercentChange < GreatestDecrease Then
                                    ActiveWorkbook.Worksheets(1).Cells(GreatestDecreaseRow, GreatestTickerCol).Value = StockTicker
                                    ActiveWorkbook.Worksheets(1).Cells(GreatestDecreaseRow, GreatestValueCol).Value = YearPercentChange
                                    GreatestDecrease = YearPercentChange
                                    If YearPercentChange < OverallGreatestDecrease Then
                                        ActiveWorkbook.Worksheets(1).Cells(GreatestDecreaseRow, GreatestOverallTickerCol).Value = StockTicker
                                        ActiveWorkbook.Worksheets(1).Cells(GreatestDecreaseRow, GreatestOverallValueCol).Value = YearPercentChange
                                        OverallGreatestDecrease = YearPercentChange
                                    End If 'Overall greatest decrease
                                End If 'Greatest decrease
                            End If 'Greatest increase
                        End If 'YearOpen > 0
                        'Analyze Volume
                        If TotalVolume > GreatestVolume Then
                            ActiveWorkbook.Worksheets(1).Cells(GreatestVolumeRow, GreatestTickerCol).Value = StockTicker
                            ActiveWorkbook.Worksheets(1).Cells(GreatestVolumeRow, GreatestValueCol).Value = TotalVolume
                            GreatestVolume = TotalVolume
                            If TotalVolume > OverallGreatestVolume Then
                                ActiveWorkbook.Worksheets(1).Cells(GreatestVolumeRow, GreatestOverallTickerCol).Value = StockTicker
                                ActiveWorkbook.Worksheets(1).Cells(GreatestVolumeRow, GreatestOverallValueCol).Value = TotalVolume
                                OverallGreatestVolume = TotalVolume
                            End If 'Overall greatest volume
                        End If 'Greatest volume
                        'Reset annual statistics and ticker state
                        FirstOpenFound = False
                        YearOpen = 0
                        YearClose = 0
                        TotalVolume = 0
                        TickerRow = TickerRow + 1
                    End If 'Last row
                End If 'First row
                'Increment ticker day row
                OriginalRowNum = OriginalRowNum + 1
            Loop 'OriginalRowNum
            'Clear debug display
            ActiveWorkbook.Worksheets(1).Cells(CurrentTickerRow, GreatestTickerCol).Value = ""
        End With 'ws
    Next SheetNum
    
    MsgBox ("Analysis Complete")
    

End Sub
