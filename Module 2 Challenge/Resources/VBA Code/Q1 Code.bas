Attribute VB_Name = "Module4"
Sub CalculateQuarterlResults()
    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim lastRow As Long, outputRow As Long
    Dim ticker As String
    Dim currentTicker As String
    Dim startDate As Date, endDate As Date
    Dim openPrice As Double, closePrice As Double
    Dim totalVolume As Double
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Q1")
    Set newWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    newWs.Name = "Quarter 1 Results"
    
    ' Add headers to the new sheet
    With newWs
        .Cells(1, 1).Value = "Ticker"
        .Cells(1, 2).Value = "Quarter"
        .Cells(1, 3).Value = "Quarterly Change"
        .Cells(1, 4).Value = "Percentage Change"
        .Cells(1, 5).Value = "Total Volume"
    End With
    
    ' Sort the data by ticker and date
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Range("A2:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ws.Sort.SortFields.Add Key:=ws.Range("B2:B" & ws.Cells(ws.Rows.Count, "B").End(xlUp).Row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange ws.Range("A1:G" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Initialize variables
    outputRow = 2
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    currentTicker = ws.Cells(2, 1).Value
    startDate = ws.Cells(2, 2).Value
    openPrice = ws.Cells(2, 3).Value
    totalVolume = 0
    
    ' Loop through each row to calculate quarterly metrics
    For i = 2 To lastRow
        ticker = ws.Cells(i, 1).Value
        If ticker <> currentTicker Then
            ' Output the results for the previous ticker
            newWs.Cells(outputRow, 1).Value = currentTicker
            newWs.Cells(outputRow, 2).Value = Year(startDate) & " Q" & Format((Month(startDate) - 1) \ 3 + 1, "0")
            newWs.Cells(outputRow, 3).Value = closePrice - openPrice
            newWs.Cells(outputRow, 4).Value = ((closePrice - openPrice) / openPrice) * 100
            newWs.Cells(outputRow, 5).Value = totalVolume
            outputRow = outputRow + 1
            currentTicker = ticker
            startDate = ws.Cells(i, 2).Value
            openPrice = ws.Cells(i, 3).Value
            totalVolume = ws.Cells(i, 7).Value
        Else
            If Year(ws.Cells(i, 2).Value) <> Year(startDate) Or (Month(ws.Cells(i, 2).Value) - 1) \ 3 <> (Month(startDate) - 1) \ 3 Then
                ' Output the results for the current quarter
                newWs.Cells(outputRow, 1).Value = currentTicker
                newWs.Cells(outputRow, 2).Value = Year(startDate) & " Q" & Format((Month(startDate) - 1) \ 3 + 1, "0")
                newWs.Cells(outputRow, 3).Value = closePrice - openPrice
                newWs.Cells(outputRow, 4).Value = ((closePrice - openPrice) / openPrice) * 100
                newWs.Cells(outputRow, 5).Value = totalVolume
                outputRow = outputRow + 1
                startDate = ws.Cells(i, 2).Value
                openPrice = ws.Cells(i, 3).Value
                totalVolume = 0
            End If
            totalVolume = totalVolume + ws.Cells(i, 7).Value
        End If
        closePrice = ws.Cells(i, 6).Value
    Next i
    
    ' Output the results for the last quarter
    newWs.Cells(outputRow, 1).Value = currentTicker
    newWs.Cells(outputRow, 2).Value = Year(startDate) & " Q" & Format((Month(startDate) - 1) \ 3 + 1, "0")
    newWs.Cells(outputRow, 3).Value = closePrice - openPrice
    newWs.Cells(outputRow, 4).Value = ((closePrice - openPrice) / openPrice) * 100
    newWs.Cells(outputRow, 5).Value = totalVolume
End Sub
