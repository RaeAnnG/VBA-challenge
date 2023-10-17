VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub MultipleYearStockData()
'Loop through all the sheets / tabs
Dim ws As Worksheet
Dim WorksheetName As String

'Create variables
Dim ticker As String
Dim NextTicker As String
Dim OpenPrice As Double
Dim highvalue As Double
Dim lowvalue As Double
Dim ClosePrice As Double
Dim TotalStockVolume As Double
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestTotal As Double

Dim LastRow As Long
Dim SummaryRow As Long

Dim i As Long
Dim k As Long
Dim Total As Long
Dim Change As Double
Dim j As Integer
Dim start As Long
Dim RowCount As Long
Dim days As Integer
Dim AverageChange As Double

Dim Sheet1(2018) As String
Dim Sheet2(2019) As String
Dim Sheet3(2020) As String

'Insert the Ticker
'Create variables to Hold Ticker, Yearly Change, Percent Change, Total Stock Volume
Dim DailyChange As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim GreatestIncreaseTicker As String
Dim GreatestDecreaseTicker As String
Dim GreatestTotalTicker As String

'Initialize TotalStockVolume for the first ticker
TotalStockVolume = 0
days = 0
years = 0

'Loop through all the Worksheets
For Each ws In Worksheets
TotalStockVolume = 0
days = 0
years = 0

'Determine the Last Row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Add a Column for the New Header Rows
ws.Range("J1").EntireColumn.Insert

'Establish the new Headers - Ticker, Yearly Change, Percent Change, Total Stock Volume
ws.Cells(1, 10).Value = "Ticker"
ws.Cells(1, 11).Value = "Yearly Change"
ws.Cells(1, 12).Value = "Percent Change"
ws.Cells(1, 13).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

'Keep track of the Location for each ticker in the Summary Table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'Initialization value - Next ticker
start = 2
previous_i = 1
OpenPrice = ws.Cells(2, 3).Value

'Loop through all the data

    For i = 2 To LastRow
    ws.Cells(i, 12).NumberFormat = "0.00%"
            
    TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
   
        'Check to see if we are still in the same ticker
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        'Grab the ticker name, opening price, closing price, total volume
        ticker = ws.Cells(i, 1).Value
        
        previous_i = previous_i + 1
        ClosePrice = ws.Cells(i, 6).Value
        YearlyChange = ClosePrice - OpenPrice
                  
        'Start the next ticker
        start = i + 1
                       
        'Calculate Percent Change
        PercentChange = YearlyChange / OpenPrice
        ws.Range("L" & Summary_Table_Row).Value = PercentChange
        ws.Cells(i, 12).Value = Format(PercentChange, "Percent")
        'ws.Cells(i, 12).Value = FormatPercent(Range("L"))
        
        'If ws.Range("L" & Summary_Table_Row).Value = PercentChange = " " Then
            'Exit Sub
        'End If
        
            'Add Color Coding to YearlyChange
            'For i = 2 To LastRow
            If YearlyChange > 0 Then
            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
            ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
            
            ElseIf YearlyChange <= 0 Then
            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
            ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
      
        OpenPrice = ws.Cells(i + 1, 3).Value
                             
        'Calculate Total Stock Volume
        'TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

        'Print the ticker, price change and total volume in the summary table
        ws.Range("J" & Summary_Table_Row).Value = ticker
        ws.Range("K" & Summary_Table_Row).Value = YearlyChange
        ws.Range("L" & Summary_Table_Row).Value = PercentChange
        ws.Range("M" & Summary_Table_Row).Value = TotalStockVolume

        Summary_Table_Row = Summary_Table_Row + 1

        ticker = 0
        ClosePrice = 0
        TotalStockVolume = 0
        YearlyChange = 0
        days = 0
        years = 0
        j = j + 1
                       
       'Add Headers for Greatest Increase, Decrease and Total
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
                         
        End If
        
Next i

        'Greatest Increase, Decrease and Total
        'Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
              
        'Find Greatest Increase, Decrease and Total
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestTotal = 0
        GreatestIncreaseTicker = ""
        GreatestDecreaseTicker = ""
        GreatestTotalTicker = ""
        
        For i = 2 To LastRow
        
        If ws.Cells(i, 12).Value >= GreatestIncrease Then
        GreatestIncreaseTicker = ws.Cells(i, 10).Value
        GreatestIncrease = ws.Cells(i, 12).Value
        
        End If
        
        If ws.Cells(i, 12).Value <= GreatestDecrease Then
        GreatestDecreaseTicker = ws.Cells(i, 10).Value
        GreatestDecrease = ws.Cells(i, 12).Value
        
        End If
        
        If ws.Cells(i, 13).Value > GreatestTotal Then
        GreatestTotalTicker = ws.Cells(i, 10).Value
        GreatestTotal = ws.Cells(i, 13).Value
        
        End If
        
        Next i
        
        'Place Values in Greatest Increase, Decrease, Total
        ws.Cells(2, 16).Value = GreatestIncreaseTicker
        ws.Cells(2, 17).Value = GreatestIncrease
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = GreatestDecreaseTicker
        ws.Cells(3, 17).Value = GreatestDecrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 16).Value = GreatestTotalTicker
        ws.Cells(4, 17).Value = GreatestTotal
        ws.Cells(4, 17).NumberFormat = "#"

ws.Range("A:Q").EntireColumn.AutoFit

Next ws

End Sub


