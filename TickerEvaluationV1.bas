Attribute VB_Name = "Module1"
Sub TickerEvaluationV1()

' Iterate through each Worksheet in the Workbook
For Each ws In Worksheets
    
' Establish Variables
Dim i As Double
Dim RecordID As Double
Dim NextRecordIDRow As Double
Dim TickerSymbol As String
Dim NextTickerSymbol As String
Dim DateStamp As String
Dim DailyStart As Currency
Dim DailyHigh As Currency
Dim DailyLow As Currency
Dim DailyEnd As Currency
Dim DailyVolume As Long
Dim YearlyStart As Currency
Dim YearlyEnd As Currency
Dim YearlyRunningHigh As Currency
Dim YearlyRunningLow As Currency
Dim YearlyChangeAmount As Currency
Dim YearlyPercentChange As Double
Dim YearlyAggregateVolume As LongLong

'Worksheet Related Variables
Dim LastRow As Double


'Variables related to Aggregate Data iteration
Dim j As Integer
j = 2


' Set initial variable values
YearlyStart = 0
YearlyEnd = 0
YearlyRunningHigh = 0
YearlyRunningLow = 0
YearlyChangeAmount = 0
YearlyPercentChange = 0
YearlyAggregateVolume = 0

ws.Columns("J").ColumnWidth = 13
ws.Columns("K").ColumnWidth = 13
ws.Columns("L").ColumnWidth = 15

'Set Header Cell Description Values for new columns
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

ws.Columns("O").ColumnWidth = 20

'Set Cell Description values for identified "Greatest" tickers
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

'Identify last row in the given worksheet
LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row


        'Initial Message Box for Start of Script
        MsgBox ("STARTING UPDATE SCRIPT" & Chr(13) & Chr(10) & _
        "Total Rows: " + Str(LastRow))



        'FOR TROUBLESHOOTING: Display values for initial variable review
        'MsgBox ("STARTING VBA SCRIPT" & Chr(13) & Chr(10) & _
        '"Yearly Start Price: " + Str(YearlyStart) & Chr(13) & Chr(10) & _
        '"Yearly Running High: " + Str(YearlyRunningHigh) & Chr(13) & Chr(10) & _
        '"Yearly Running Low: " + Str(YearlyRunningLow) & Chr(13) & Chr(10) & _
        '"Yearly Ending Price" + Str(YearlyEnd) & Chr(13) & Chr(10) & _
        '"Yearly Change Amount: " + Str(YearlyChangeAmount) & Chr(13) & Chr(10) & _
        '"Yearly Percent Change: " + Str(YearlyPercentChange) & Chr(13) & Chr(10) & _
        '"Yearly Aggregate Volume: " + Str(YearlyAggregateVolume) & Chr(13) & Chr(10) & _
        '" " & Chr(13) & Chr(10) & _
        '"Total Rows: " + Str(LastRow))


    ' Create a for loop from 1 to N to evaluate each row of data
    For i = 2 To LastRow    ' 525 - First Two Tickers
    
    
        'Pull back cell values for row-by-row evaluation
        RecordID = (i - 1)
        NextRecordIDRow = (i + 1)
        TickerSymbol = ws.Cells(i, 1).Value
        NextTickerSymbol = ws.Cells(NextRecordIDRow, 1).Value
        DateStamp = ws.Cells(i, 2).Value
        DailyStart = ws.Cells(i, 3).Value
        DailyHigh = ws.Cells(i, 4).Value
        DailyLow = ws.Cells(i, 5).Value
        DailyEnd = ws.Cells(i, 6).Value
        DailyVolume = ws.Cells(i, 7).Value
        
        ' Evaluate/Populate Yearly Start value
        If YearlyStart = 0 Then YearlyStart = DailyStart
        
        ' Evaluate/Populate Yearly Running High value
        If YearlyRunningHigh = 0 Then
           YearlyRunningHigh = DailyHigh
        Else
        If (YearlyRunningHigh <> 0 And YearlyRunningHigh > DailyHigh) Then
            YearlyRunningHigh = YearlyRunningHigh
        Else
        If (YearlyRunningHigh <> 0 And YearlyRunningHigh < DailyHigh) Then
            YearlyRunningHigh = DailyHigh
        End If
        End If
        End If
        
        ' Evaluate/Populate Yearly Running Low value
        If YearlyRunningLow = 0 Then
           YearlyRunningLow = DailyLow
        Else
        If (YearlyRunningLow <> 0 And YearlyRunningLow < DailyLow) Then
        YearlyRunningLow = YearlyRunningLow
        Else
        If (YearlyRunningLow <> 0 And YearlyRunningLow > DailyLow) Then
        YearlyRunningLow = DailyLow
        End If
        End If
        End If
        
        ' Evaluate/Populate Yearly Aggregate Volume value
        If YearlyAggregateVolume = 0 Then
           YearlyAggregateVolume = DailyVolume
        Else
        If YearlyAggregateVolume <> 0 Then
           YearlyAggregateVolume = (YearlyAggregateVolume + DailyVolume)
        End If
        End If
        
        ' Evaluate/Populate Yearly Ending Price value
        If TickerSymbol = NextTickerSymbol Then
           YearlyEnd = 0
        Else
        If TickerSymbol <> NextTickerSymbol Then
           YearlyEnd = DailyEnd
        End If
        End If
        
        
         ' Evaluate/Populate Yearly Change Amount value
        If TickerSymbol = NextTickerSymbol Then
           YearlyChangeAmount = 0
        Else
        If TickerSymbol <> NextTickerSymbol Then
           YearlyChangeAmount = (YearlyEnd - YearlyStart)
        End If
        End If
        
        
        
        ' Evaluate/Populate Yearly Percentage Change value
        If TickerSymbol = NextTickerSymbol Then
           YearlyPercentChange = 0
        Else
        If (TickerSymbol <> NextTickerSymbol And YearlyStart = 0 And YearlyChangeAmount = 0) Then
           YearlyPercentChange = 0
        Else
        If (TickerSymbol <> NextTickerSymbol And YearlyStart <> 0 And YearlyChangeAmount <> 0) Then
           YearlyPercentChange = ((YearlyChangeAmount) / YearlyStart)
        End If
        End If
        End If
        

        
        'Populate New Fields with aggregated Data for the given Ticker
        If TickerSymbol <> NextTickerSymbol Then
        ws.Cells(j, 9).Value = TickerSymbol
        ws.Cells(j, 10).Value = Str(YearlyChangeAmount)
        ws.Cells(j, 11).Value = Str(YearlyPercentChange)
        ws.Cells(j, 12).Value = Str(YearlyAggregateVolume)
        End If
        
        'If the value is greater than or equal to zero = set background color equal to green
        If YearlyChangeAmount >= 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 4
        End If
        
        'If the value is less than zero = set background color equal to red
        If YearlyChangeAmount < 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 3
        End If
        
        'Format Value Yearly Change as Currency
         ws.Cells(j, 10).NumberFormat = "$#,##0.00"
        
        'Format Value Percent Change as Percent
         ws.Cells(j, 11).NumberFormat = "0.00%"
                         
        'Format Value column as Percent
         ws.Cells(2, 15).NumberFormat = "0.00%"
         ws.Cells(3, 15).NumberFormat = "0.00%"
         
   
        ' Attempted to get the Min/Max Percentages for Bonus- but this was evaluating on each row - boggging down. Decided to punt.
        'Dim PercentRange As Range
        'LastSummaryRow = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row
        'Identify the smallest Percent Change
        'Set PercentRange = ws.Range("K2:K300")
        'Worksheet function MIN returns the smallest value in a range
        'MinPercentChange = Application.WorksheetFunction.Min(PercentRange)
        'MaxPercentChange = Application.WorksheetFunction.Max(PercentRange)
        'ws.Cells(2, 15).Value = Str(MaxPercentChange)
        'ws.Cells(3, 15).Value = Str(MinPercentChange)
        
        
        'FOR INDIVIDUAL TROUBLESHOOTING: Display values for iterative variable review
        'MsgBox ("Record ID: " + Str(RecordID) & Chr(13) & Chr(10) & _
        '"Ticker Symbol: " + TickerSymbol & Chr(13) & Chr(10) & _
        '"Date Stamp: " + DateStamp & Chr(13) & Chr(10) & _
        '"Daily Start: " + Str(DailyStart) & Chr(13) & Chr(10) & _
        '"Daily High: " + Str(DailyHigh) & Chr(13) & Chr(10) & _
        '"Daily Low: " + Str(DailyLow) & Chr(13) & Chr(10) & _
        '"Daily End: " + Str(DailyEnd) & Chr(13) & Chr(10) & _
        '" " & Chr(13) & Chr(10) & _
        '"YEARLY RUNNING VALUES " & Chr(13) & Chr(10) & _
        '"Yearly Start Price: " + Str(YearlyStart) & Chr(13) & Chr(10) & _
        '"Yearly Running High: " + Str(YearlyRunningHigh) & Chr(13) & Chr(10) & _
        '"Yearly Running Low: " + Str(YearlyRunningLow) & Chr(13) & Chr(10) & _
        '"Yearly Ending Price" + Str(YearlyEnd) & Chr(13) & Chr(10) & _
        '"Yearly Change Amount: " + Str(YearlyChangeAmount) & Chr(13) & Chr(10) & _
        '"Yearly Percent Change: " + Str(YearlyPercentChange) & Chr(13) & Chr(10) & _
        '"Yearly Aggregate Volume: " + Str(YearlyAggregateVolume))
       
        'FOR TICKER AGGREGATE TROUBLESHOOTING: Display values for Each Ticker Aggregation review
        'If TickerSymbol <> NextTickerSymbol Then
        'MsgBox ("Record ID: " + Str(RecordID) & Chr(13) & Chr(10) & _
        '"Ticker Symbol: " + TickerSymbol & Chr(13) & Chr(10) & _
        '"Date Stamp: " + DateStamp & Chr(13) & Chr(10) & _
        '"Daily Start: " + Str(DailyStart) & Chr(13) & Chr(10) & _
        '"Daily High: " + Str(DailyHigh) & Chr(13) & Chr(10) & _
        '"Daily Low: " + Str(DailyLow) & Chr(13) & Chr(10) & _
        '"Daily End: " + Str(DailyEnd) & Chr(13) & Chr(10) & _
        '" " & Chr(13) & Chr(10) & _
        '"YEARLY RUNNING VALUES " & Chr(13) & Chr(10) & _
        '"Yearly Start Price: " + Str(YearlyStart) & Chr(13) & Chr(10) & _
        '"Yearly Running High: " + Str(YearlyRunningHigh) & Chr(13) & Chr(10) & _
        '"Yearly Running Low: " + Str(YearlyRunningLow) & Chr(13) & Chr(10) & _
        '"Yearly Ending Price" + Str(YearlyEnd) & Chr(13) & Chr(10) & _
        '"Yearly Change Amount: " + Str(YearlyChangeAmount) & Chr(13) & Chr(10) & _
        '"Yearly Percent Change: " + Str(YearlyPercentChange) & Chr(13) & Chr(10) & _
        '"Yearly Aggregate Volume: " + Str(YearlyAggregateVolume))
        ' End If
       
       'Re-setting Variables to Zero for next Iterating through next Ticker Value
        If TickerSymbol <> NextTickerSymbol Then
            TickerSymbol = ""
            DateStamp = ""
            DailyStart = 0
            DailyHigh = 0
            DailyLow = 0
            DailyEnd = 0
            YearlyStart = 0
            YearlyEnd = 0
            YearlyRunningHigh = 0
            YearlyRunningLow = 0
            YearlyChangeAmount = 0
            YearlyPercentChange = 0
            YearlyAggregateVolume = 0
            j = (j + 1)
         End If

    Next i
    
        'FOR TROUBLESHOOTING: Display FINAL values for review
        'MsgBox ("FINAL YEARLY VALUES " & Chr(13) & Chr(10) & _
        '"Ticker Symbol: " + TickerSymbol & Chr(13) & Chr(10) & _
        '"Yearly Start Price: " + Str(YearlyStart) & Chr(13) & Chr(10) & _
        '"Yearly Running High: " + Str(YearlyRunningHigh) & Chr(13) & Chr(10) & _
        '"Yearly Running Low: " + Str(YearlyRunningLow) & Chr(13) & Chr(10) & _
        '"Yearly Ending Price" + Str(YearlyEnd) & Chr(13) & Chr(10) & _
        '"Yearly Change Amount: " + Str(YearlyChangeAmount) & Chr(13) & Chr(10) & _
        '"Yearly Percent Change: " + Str(YearlyPercentChange) & Chr(13) & Chr(10) & _
        '"Yearly Aggregate Volume: " + Str(YearlyAggregateVolume))



Next ws

        'Initial Message Box for Start of Script
        MsgBox ("END UPDATE SCRIPT" & Chr(13) & Chr(10) & _
        "Total Rows: " + Str(LastRow))

End Sub

