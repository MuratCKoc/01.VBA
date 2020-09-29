Sub TickerCounter()

    
    ' Main variable Declarations
    Dim tickerName As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim greatestIncrease, greatestDecrease As Double
    Dim greatestVolume As LongLong
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalStockVolume As LongLong
    Dim tableRow, lastRow, lastColumn As Integer
    Dim currentSheet As Worksheet
    
    
    ' Worksheet Loop
    For Each currentSheet In Worksheets
    
    ' Initialize & set counters
    tableRow = 2
    openPrice = 0
    lastRow = 0
    lastRow = 0
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    
    ' Add row headers
    currentSheet.Range("I1").Value = "Ticker"
    currentSheet.Range("J1").Value = "Yearly Change"
    currentSheet.Range("K1").Value = "Percent Change"
    currentSheet.Range("L1").Value = "Total Stock Volume"
    currentSheet.Range("O2").Value = "Greatest % Increase"
    currentSheet.Range("O3").Value = "Greatest % Decrease"
    currentSheet.Range("O4").Value = "Greatest Total Volume"
    currentSheet.Range("P1").Value = "Ticker"
    currentSheet.Range("Q1").Value = "Value"
    
    lastRow = currentSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop thru all stock prices
    For i = 2 To lastRow
        
        ' Pick the opening Price
        If IsEmpty(openPrice) Or openPrice = 0 Then
            openPrice = currentSheet.Cells(i, 3)
        End If
        
        ' Check if we are within the same ticker
        If currentSheet.Cells(i + 1, 1).Value <> currentSheet.Cells(i, 1).Value Then
            
            ' Set Ticker name
            tickerName = currentSheet.Cells(i, 1).Value
            
            ' Set closing Price
            closePrice = currentSheet.Cells(i, 6).Value
            
            ' Print Ticker
            currentSheet.Range("I" & tableRow).Value = tickerName
            
            ' Set and Print yearly change
            yearlyChange = closePrice - openPrice
            currentSheet.Range("J" & tableRow).Value = closePrice - openPrice
            
            ' Red & Green background color indicator of status
            If yearlyChange < 0 Then
            currentSheet.Range("J" & tableRow).Interior.ColorIndex = 3
            Else
            currentSheet.Range("J" & tableRow).Interior.ColorIndex = 4
            End If
            
            ' Set and Print percentage change
            If yearlyChange = 0 Then
                percentChange = 0
            Else
                percentChange = (yearlyChange / openPrice) * 100
            End If
            currentSheet.Range("K" & tableRow).Value = percentChange
            
            ' Print Total Stock Volume
            totalStockVolume = totalStockVolume + currentSheet.Cells(i, 7).Value
            currentSheet.Range("L" & tableRow).Value = totalStockVolume
            
            ' Increment and Reset counters
            tableRow = tableRow + 1
            totalStockVolume = 0
            openPrice = 0
            closePrice = 0
            percentChange = 0
            
        Else
            ' Keep adding total if same ticker
            totalStockVolume = totalStockVolume + currentSheet.Cells(i, 7).Value
        End If
           
    
    Next i
    ' End Sheet Loop
    
    ' Greatest Increase
    greatestIncrease = WorksheetFunction.Max(currentSheet.Range("K2:K" & lastRow))
    tickerName = WorksheetFunction.Match(greatestIncrease, currentSheet.Range("K2:K" & lastRow), 0)
    currentSheet.Range("P2").Value = currentSheet.Range("I" & tickerName + 1)
    currentSheet.Range("Q2").Value = greatestIncrease

    ' Greatest Decrease
    greatestDecrease = WorksheetFunction.Min(currentSheet.Range("K2:K" & lastRow))
    tickerName = WorksheetFunction.Match(greatestDecrease, currentSheet.Range("K2:K" & lastRow), 0)
    currentSheet.Range("P3").Value = currentSheet.Range("I" & tickerName + 1)
    currentSheet.Range("Q3").Value = greatestDecrease

    ' Greatest Stock Volume
    totalStockVolume = WorksheetFunction.Max(currentSheet.Range("L2:L" & lastRow))
    tickerName = WorksheetFunction.Match(totalStockVolume, currentSheet.Range("L2:L" & lastRow), 0)
    currentSheet.Range("P4").Value = currentSheet.Range("I" & tickerName + 1)
    currentSheet.Range("Q4").Value = totalStockVolume

    ' Cell Percentage Formatting
    currentSheet.Range("K2:K" & lastRow).NumberFormat = "0.00\%"
    currentSheet.Range("Q2:Q3").NumberFormat = "0.00\%"
    currentSheet.Range("A:Q").Columns.AutoFit
    
    ' Cell Percentage Formatting
    currentSheet.Range("K2:K" & lastRow).NumberFormat = "0.00\%"
    currentSheet.Range("Q2:Q3").NumberFormat = "0.00\%"
    currentSheet.Range("A:Q").Columns.AutoFit
Next
' End of Worksheets Loop
End Sub

