Attribute VB_Name = "Module1"
Sub Stock()

Dim totalVol As Double
Dim openingPrice As Single
Dim closingPrice As Single
Dim percentageChange As Single
Dim resultList As Integer
Dim priceDifference As Single
Dim wsCount As Integer


'FOR EACH WORKSHEET
wsCount = ActiveWorkbook.Worksheets.Count
For wsi = 1 To wsCount

    Worksheets(wsi).Activate

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row


'HEADER STORAGE
    Dim headers As Variant
    headers = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")

'POPULATE HEADERS
    For i = 0 To 3
        Cells(1, i + 9).Value = headers(i)
        Next i

'SET BASE VALUES FOR POPULATING TABLE 2
    resultList = 2
    totalVol = 0

'POPULATE TABLE 2
    For i = 2 To lastRow
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            closingPrice = Cells(i, 6).Value
            totalVol = totalVol + Cells(i, 7).Value
            Cells(resultList, 9).Value = Cells(i, 1).Value
            Cells(resultList, 12).Value = totalVol
            priceDifference = closingPrice - openingPrice
            Cells(resultList, 10).NumberFormat = "0.00"
            Cells(resultList, 10).Value = priceDifference
        
'COLOR YEARLY CHANGE BASED ON POSITIVE OR NEGATIVE VALUE
            If priceDifference < 0 Then
                Cells(resultList, 10).Interior.ColorIndex = 3
            Else
                Cells(resultList, 10).Interior.ColorIndex = 4
            End If
        
'CALCULATE AND RECORD PERCENT CHANGE
            Cells(resultList, 11).Value = FormatPercent(priceDifference / openingPrice)
        
'COLOR PERCENT CHANGE BASED ON POSITIVE OR NEGATIVE VALUE
            If Cells(resultList, 11).Value < 0 Then
                Cells(resultList, 11).Interior.ColorIndex = 3
            Else
                Cells(resultList, 11).Interior.ColorIndex = 4
            End If
            
'RESET TOTALVOL AND INCREMENT RESULTLIST
            totalVol = 0
            resultList = resultList + 1
        
        ElseIf (Cells(i, 1).Value = Cells(i + 1, 1).Value) And (Cells(i, 1).Value <> Cells(i - 1, 1).Value) Then
            openingPrice = Cells(i, 3).Value
            totalVol = totalVol + Cells(i, 7).Value
        Else
            totalVol = totalVol + Cells(i, 7).Value
        End If
    Next i


'RECALCULATE LASTROW FOR SECOND TABLE
    lastRow = Cells(Rows.Count, 9).End(xlUp).Row

'VARIABLES FOR TABLE 3
    Dim greatestIncreaseTicker As String, greatestDecreaseTicker As String, greatestTotalVolTicker As String
    Dim greatestIncreasePercent As Single, greatestDecreasePercent As Single, greatestTotalVol As Double

'ROW HEADER STORAGE FOR TABLE 3
    headers = Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume")

'POPULATE ROW HEADERS FOR TABLE 3
    For i = 0 To 2
        Cells(2 + i, 15).Value = headers(i)
        Next i
    
'COLUMN HEADER STORAGE FOR TABLE 3
    headers = Array("Ticker", "Value")

'POPULATE COLUMN HEADERS FOR TABLE 3
    For i = 0 To 1
        Cells(1, 16 + i).Value = headers(i)
        Next i

'FIND TABLE 3 VALUES
    greatestIncreasePercent = Cells(i, 11).Value
    greatestDecreasePercent = Cells(i, 11).Value
    greatestTotalVol = Cells(i, 12).Value
    greatestIncreaseTicker = Cells(i, 10).Value
    greatestDecreaseTicker = Cells(i, 10).Value
    greatestTotalVolTicker = Cells(i, 10).Value

    For i = 2 To lastRow
        If Cells(i, 11).Value > greatestIncreasePercent Then
            greatestIncreasePercent = Cells(i, 11).Value
            greatestIncreaseTicker = Cells(i, 9).Value
        End If
        If Cells(i, 11).Value < greatestDecreasePercent Then
            greatestDecreasePercent = Cells(i, 11).Value
            greatestDecreaseTicker = Cells(i, 9).Value
        End If
        If Cells(i, 12).Value > greatestTotalVol Then
            greatestTotalVol = Cells(i, 12).Value
            greatestTotalVolTicker = Cells(i, 9).Value
        End If
    Next i
    
'POPULATE TABLE 3 VALUES
    Cells(2, 16).Value = greatestIncreaseTicker
    Cells(2, 17).Value = FormatPercent(greatestIncreasePercent)
    Cells(3, 16).Value = greatestDecreaseTicker
    Cells(3, 17).Value = FormatPercent(greatestDecreasePercent)
    Cells(4, 16).Value = greatestTotalVolTicker
    Cells(4, 17).Value = greatestTotalVol


'AUTOFIT TO DISPLAY DATA
    Columns("A:Q").AutoFit


Next wsi

End Sub
