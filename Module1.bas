Attribute VB_Name = "Module1"
Sub easyVolCount()
Call volCount(0)
End Sub

Sub modVolCount()
Call volCount(1)
End Sub

Sub hardVolCount()
Call volCount(2)
End Sub

Sub chalVolCount()
Call volCount(3)
End Sub

Sub volCount(skill)
Dim currsheet As Worksheet
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
If skill = 3 Then
    'Loop through all sheets
    startsheet = 1
    endsheet = ActiveWorkbook.Worksheets.Count
Else
    'Only loop current sheet
    startsheet = ActiveSheet.Index
    endsheet = ActiveSheet.Index
End If
For sheetindex = startsheet To endsheet
    Worksheets(sheetindex).Activate
    Range("I:Q").ClearContents
    Range("I:Q").ClearFormats
    entcnt = Application.WorksheetFunction.CountA(Range("A:A"))
    Dim stocks() As Variant
    stocks = Range(Cells(1, 1), Cells(entcnt + 1, 7)).Value
    'Label sheet
    If skill = 0 Then
        volCol = 10
    Else
        volCol = 12
        Range("J1") = "Yearly Change"
        Range("K1") = "Percent Change"
        If skill > 1 Then
            Range("O2") = "Greatest % Increase"
            Range("O3") = "Greatest % Decrease"
            Range("O4") = "Greatest Total Volume"
            Range("P1") = "Ticker"
            Range("Q1") = "Value"
        End If
    End If
    Range("I1") = "Ticker"
    Cells(1, volCol) = "Total Stock Volume"
    'Initialize first copmany
    tckr = 1
    currtot = 0
    mostTot = 0
    mostGrow = 0
    mostLose = 0
    yrOpen = stocks(2, 3)
    'Scan through all entries
    For curr = 2 To entcnt
        currtick = stocks(curr, 1)
        currtot = stocks(curr, 7) + currtot
        'Check if last tick for company
        If currtick <> stocks(curr + 1, 1) Then
            'Close out year
            If skill <> 0 Then
                yrClose = stocks(curr, 6)
                Cells(tckr + 1, 10) = yrClose - yrOpen
                If yrOpen > yrClose Then
                    Cells(tckr + 1, 10).Interior.Color = RGB(255, 0, 0)
                Else
                    Cells(tckr + 1, 10).Interior.Color = RGB(0, 255, 0)
                End If
                If yrOpen <> 0 Then
                    grow = (yrClose - yrOpen) / yrOpen
                Else
                    grow = 0
                End If
                Cells(tckr + 1, 11) = grow
                'Check superlatives
                If skill > 1 Then
                    If grow > mostGrow Then
                        Range("P2") = currtick
                        mostGrow = grow
                    ElseIf grow < mostLose Then
                        Range("P3") = currtick
                        mostLose = grow
                    End If
                    If currtot > mostTot Then
                        Range("P4") = currtick
                        mostTot = currtot
                    End If
                End If
            End If
            'Print company data
            Cells(tckr + 1, 9) = currtick
            Cells(tckr + 1, volCol) = currtot
            'Advance to new company
            currtot = 0
            tckr = tckr + 1
            yrOpen = stocks(curr + 1, 3)
        End If
    Next curr
    'Format Percent and print superlatives
    If skill <> 0 Then
        Range(Cells(2, 11), Cells(tckr + 1, 11)).NumberFormat = "0.00%"
        If skill > 1 Then
            Range("Q2") = mostGrow
            Range("Q3") = mostLose
            Range("Q2:Q3").NumberFormat = "0.00%"
            Range("Q4") = mostTot
        End If
    End If
    Columns("I:Q").AutoFit
Next sheetindex
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub


