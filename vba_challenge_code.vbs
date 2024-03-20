Sub ticker_analysis():

    ' Define variables and dimensions
    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim b As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim days As Integer
    Dim dailyChange As Double
    Dim averageChange As Double

    ' Define title row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

    ' Define initial values
    b = 0
    total = 0
    change = 0
    start = 2

    ' RowCount to determind total number of rows
    rowCount = Cells(Rows.Count, "A").End(xlUp).Row

    For i = 2 To rowCount

        ' Change in ticker values recorded here
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' Results stored in total variable
            total = total + Cells(i, 7).Value

            ' Deifne instance of toal value equal to zero
            If total = 0 Then
                ' print the results
                Range("I" & 2 + b).Value = Cells(i, 1).Value
                Range("J" & 2 + b).Value = 0
                Range("K" & 2 + b).Value = "%" & 0
                Range("L" & 2 + b).Value = 0

            Else
                ' Define new value 
                If Cells(start, 3) = 0 Then
                    For new_value = start To i
                        If Cells(new_value, 3).Value <> 0 Then
                            start = new_value
                            Exit For
                        End If
                     Next new_value
                End If

                ' Change and percent change equation defined
                change = (Cells(i, 6) - Cells(start, 3))
                percent_change = change / Cells(start, 3)

                ' Stock ticker loop define as "start"
                start = i + 1

                ' Results are printed here
                Range("I" & 2 + b).Value = Cells(i, 1).Value
                Range("J" & 2 + b).Value = change
                Range("J" & 2 + b).NumberFormat = "0.00"
                Range("K" & 2 + because).Value = percent_change
                Range("K" & 2 + b).NumberFormat = "0.00%"
                Range("L" & 2 + b).Value = total

                ' Color code negative and positvies
                Select Case change
                    Case Is > 0
                        Range("J" & 2 + b).Interior.ColorIndex = 4
                    Case Is < 0
                        Range("J" & 2 + b).Interior.ColorIndex = 3
                    Case Else
                        Range("J" & 2 + b).Interior.ColorIndex = 0
                End Select

            End If

            ' New stock ticker is reset and defined here
            total = 0
            change = 0
            b = b + 1
            days = 0

        Else
            total = total + Cells(i, 7).Value

        End If

    Next i

    ' Begin to calculate max and min values
    Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & rowCount)) * 100
    Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & rowCount)) * 100
    Range("Q4") = WorksheetFunction.Max(Range("L2:L" & rowCount))

    ' increase, decrease and volume calcualtions made in this block
    increase = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
    decrease = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
    volume = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & rowCount)), Range("L2:L" & rowCount), 0)

    ' Define ticker symbol for increase decrease and volume
    Range("P2") = Cells(increase + 1, 9)
    Range("P3") = Cells(decrease + 1, 9)
    Range("P4") = Cells(volume + 1, 9)

End Sub
