Attribute VB_Name = "Module1"
Sub stock()
    Dim ws As Worksheet
    Dim i As Long
    Dim c As Long
    Dim opening As Double
    Dim closing As Double
    Dim row_a As Long
    Dim row_b As Long
    Dim row_count As Long
    Dim alpha_increase As Double
    Dim alpha_decrease As Double
    Dim alpha_increase_ticker As String
    Dim alpha_decrease_ticker As String
    Dim alpha_volume As Double
    Dim alpha_volume_ticker As String
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        Range("O2") = "Greatest % Increase"
        Range("O3") = "Greatest % Decrease"
        Range("O4") = "Greatest Total Volume"
        
        Range("P1") = "Ticker"
        Range("Q1") = "Value"

        'determine the row count
        Range("A1").Select
        Selection.End(xlDown).Select
        row_count = Selection.Row
        
        'set initial values for the counters and maximums
        c = 1
        Total = 0
        alpha_increase = 0
        alpha_decrease = 0
        alpha_pct_increase = 0
        alpha_pct_decrease = 0
        alpha_increase_ticker = ""
        alpha_decrease_ticker = ""
        alpha_volume = 0
        alpha_volume_ticker = ""
        
        For i = 2 To row_count
            'if its the first
            If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                c = c + 1
                'insert the new ticker symbol
                Cells(c, 9).Value = Cells(i, 1).Value
                'record the opening price
                opening = Cells(i, 3).Value
                'record the opening row index
                Cells(i, 1).Select
                row_a = Selection.Row
            ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                'store the closing price
                closing = Cells(i, 6).Value
                'determine and store the closing row index
                Cells(i, 1).Select
                row_b = Selection.Row
                'compute and input the total volume
                myRange = Range(Cells(row_a, 7), Cells(row_b, 7))
                Cells(c, 12) = WorksheetFunction.Sum(myRange)
                If Cells(c, 12).Value > alpha_volume Then
                    alpha_volume = Cells(c, 12).Value
                    alpha_volume_ticker = Cells(i, 1).Value
                End If
                'calculate and insert the change value and set the color format
                Cells(c, 10).Value = closing - opening
                If Cells(c, 10).Value >= 0 Then
                    Cells(c, 10).Interior.ColorIndex = 4
                    Cells(c, 10).Font.ColorIndex = 0
                Else
                    Cells(c, 10).Interior.ColorIndex = 3
                    Cells(c, 10).Font.ColorIndex = 0
                End If
                'calculate and insert the percent change value
                If opening = 0 Then
                    Cells(c, 11).Value = 0
                Else
                    Cells(c, 11).Value = Cells(c, 10).Value / opening
                    Cells(c, 11).Select
                    Selection.NumberFormat = "0.00%"
                End If
                If Cells(c, 10).Value > 0 And Cells(c, 10).Value > alpha_increase Then
                    alpha_increase = Cells(c, 10).Value
                    alpha_pct_increase = Cells(c, 11).Value
                    alpha_increase_ticker = Cells(i, 1).Value
                ElseIf Cells(c, 10).Value < 0 And Cells(c, 10).Value < alpha_decrease Then
                    alpha_decrease = Cells(c, 10).Value
                    alpha_pct_decrease = Cells(c, 11).Value
                    alpha_decrease_ticker = Cells(i, 1).Value
                End If
            End If
        Next i
        Range("P2").Value = alpha_increase_ticker
        Range("P3").Value = alpha_decrease_ticker
        Range("P4").Value = alpha_volume_ticker
        Range("Q2").Value = alpha_pct_increase
        Range("Q3").Value = alpha_pct_decrease
        Range("Q4").Value = alpha_volume
        Range("Q2:Q3").Select
        Selection.NumberFormat = "0.00%"
        Columns("I:Q").AutoFit
    Next
End Sub

