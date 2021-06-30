Attribute VB_Name = "Module1"
Sub StockSum()

'Loop through all worksheets
Dim WSCount As Integer
WSCount = ActiveWorkbook.Worksheets.Count

For n = 1 To WSCount
    Worksheets(n).Activate
    
    'Ticker
    'Count # of tickers
    RowNum = WorksheetFunction.CountA(Columns(1))
    TCount = 0
    For i = 2 To RowNum
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            TCount = TCount + 1
        End If
    Next i

    'Copy Unique Tickers
        'Make column A an array
        'Fill column I with for loop, which exits when the active cell gets filled
        'StopIndex tracks at which index the for loop exits
        'Use StopIndex to skip the loop to the next ticker
    Dim TickerList As Variant
    TickerList = Range(Cells(1, 1), Cells(RowNum + 1, 1))
    
    StopIndex = 0
    For i = 2 To TCount + 1
        For j = StopIndex + 2 To RowNum
            If TickerList(j, 1) <> TickerList(j + 1, 1) Then
                Cells(i, 9) = TickerList(j, 1)
                StopIndex = j
                Exit For
            End If
        Next j
    Next i
    
    'Yearly Change
    'Figure out First Opening and Last Closing values
        'Split values into two arrays respectively
    Dim FirstOpen() As Variant, LastClose() As Variant
    ReDim FirstOpen(1 To TCount, 1) As Variant, LastClose(1 To TCount, 1) As Variant
    
    'FirstOpen(1,1) is always the first entry of column C
    'LastClose(TCount,1) is always the last entry of column F
    FirstOpen(1, 1) = Cells(2, 3)
    LastClose(TCount, 1) = Cells(RowNum, 6)
    
    'Populate array and fill column J using similar technique as above, include formatting
    StopIndex = 0
    For j = 1 To TCount - 1
        For i = 1 + StopIndex To RowNum - 1
            If Cells(i + 1, 1) <> Cells(i + 2, 1) Then
                LastClose(j, 1) = Cells(i + 1, 6)
                FirstOpen(j + 1, 1) = Cells(i + 2, 3)
                    If (LastClose(j, 1) - FirstOpen(j, 1)) >= 0 Then
                        Cells(j + 1, 10) = LastClose(j, 1) - FirstOpen(j, 1)
                        Cells(j + 1, 10).Interior.ColorIndex = 4
                        Range(Cells(j + 1, 10), Cells(j + 1, 10)).NumberFormat = "+0.00;-0.00"
                    Else
                        Cells(j + 1, 10) = LastClose(j, 1) - FirstOpen(j, 1)
                        Cells(j + 1, 10).Interior.ColorIndex = 3
                        Range(Cells(j + 1, 10), Cells(j + 1, 10)).NumberFormat = "+0.00;-0.00"
                    End If
                StopIndex = i
                Exit For
            End If
        Next i
    Next j
    
    'Make sure to fill the final row of column J
    Cells(TCount + 1, 10) = LastClose(TCount, 1) - FirstOpen(TCount, 1)
    If Cells(TCount + 1, 10) >= 0 Then
        Cells(TCount + 1, 10).Interior.ColorIndex = 4
    Else
        Cells(TCount + 1, 10).Interior.ColorIndex = 3
    End If
    
    'Percent Change
    'Use FirstOpen and LastClose to find percent change
    For i = 2 To TCount + 1
        If FirstOpen(i - 1, 1) <> 0 Then
            Cells(i, 11) = (LastClose(i - 1, 1) - FirstOpen(i - 1, 1)) / FirstOpen(i - 1, 1)
       Else
            Cells(i, 11) = 0
        End If
    Next i
    
    'Stock Volume
    StopIndex = 2
    For i = 2 To TCount + 1
        For j = StopIndex To RowNum
                If TickerList(j, 1) <> TickerList(j + 1, 1) Then
                    Cells(i, 12) = WorksheetFunction.Sum(Range(Cells(StopIndex, 7), Cells(j, 7)))
                    StopIndex = j + 1
                    Exit For
                End If
        Next j
    Next i
    
    'Greatest % Increase
    Range("P2:P2") = WorksheetFunction.Max(Range(Cells(2, 11), Cells(TCount + 1, 11)))
    
    'Want to account for possible duplicates
    RepCount = 0
    For i = 2 To TCount + 1
        If Cells(i, 11) = Range("P2:P2") Then
            RepCount = RepCount + 1
        End If
    Next i
    
    Dim IncReps() As Variant
    ReDim IncReps(1 To RepCount) As Variant
    
    StopIndex = 2
    For i = 1 To RepCount
        For j = StopIndex To TCount + 1
            If Cells(j, 11) = Range("P2:P2") Then
                IncReps(i) = Cells(j, 9)
                StopIndex = j + 1
                Exit For
            End If
        Next j
    Next i
    
    If RepCount = 1 Then
        Range("O2:O2") = IncReps(1)
    ElseIf RepCount > 1 Then
        Range("O2:O2") = Join(IncReps, ", ")
    End If
    
    'Greatest % Decrease
    Range("P3:P3") = WorksheetFunction.Min(Range(Cells(2, 11), Cells(TCount + 1, 11)))
    
    RepCount = 0
    For i = 2 To TCount + 1
        If Cells(i, 11) = Range("P3:P3") Then
            RepCount = RepCount + 1
        End If
    Next i
    
    Dim DecReps() As Variant
    ReDim DecReps(1 To RepCount) As Variant
    
    StopIndex = 2
    For i = 1 To RepCount
        For j = StopIndex To TCount + 1
            If Cells(j, 11) = Range("P3:P3") Then
                DecReps(i) = Cells(j, 9)
                StopIndex = j + 1
                Exit For
            End If
        Next j
    Next i
    
    If RepCount = 1 Then
        Range("O3:O3") = DecReps(1)
    ElseIf RepCount > 1 Then
        Range("O3:O3") = Join(DecReps, ", ")
    End If
    
    'Greatest Total Volume
    Range("P4:P4") = WorksheetFunction.Max(Range(Cells(2, 12), Cells(TCount + 1, 12)))
    
    RepCount = 0
    For i = 2 To TCount + 1
        If Cells(i, 12) = Range("P4:P4") Then
            RepCount = RepCount + 1
        End If
    Next i
    
    Dim VolReps() As Variant
    ReDim VolReps(RepCount) As Variant
    
    StopIndex = 2
    For i = 1 To RepCount
        For j = StopIndex To TCount + 1
            If Cells(j, 12) = Range("P4:P4") Then
                VolReps(i) = Cells(j, 9)
                StopIndex = j + 1
                Exit For
            End If
        Next j
    Next i
    
    If RepCount = 1 Then
        Range("O4:O4") = VolReps(1)
    ElseIf RepCount > 1 Then
        Range("O4:O4") = Join(VolReps(), ", ")
    End If
    
    'Format
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    Range("N2") = "Greatest % Increase"
    Range("N3") = "Greatest % Decrease"
    Range("N4") = "Greatest Total Volume"
    Range("O1") = "Ticker"
    Range("P1") = "Value"
    Columns("A:P").AutoFit
    Range("I1:L1,N2:N4,O1:P1").Font.FontStyle = "Bold"
    Range("P2:P3").NumberFormat = "0.00%"
    Range(Cells(2, 11), Cells(TCount + 1, 11)).NumberFormat = "0.00%"
    Range(Cells(TCount + 1, 10), Cells(TCount, 10)).NumberFormat = "+0.00;-0.00"

    'Deletes the Program Output
        '''Range(Cells(1, 9), Cells(TCount + 1, 16)).Delete

Next n

End Sub

