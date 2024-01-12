Sub YearStock()
    Dim Lastrow As Long
    Dim ws As Worksheet

For Each ws In Worksheets
        Dim WorksheetName As String
        Lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
       ' WorksheetName = ws.Name
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

        ' Ticker Type
        Dim Ticker_type As String
        Dim Yearly_Change As Double
        Dim Total_Stock As Double

        Dim opeingpr As Double
        openingpr = ws.Cells(2, 3).Value

        Yearly_Change = 0
        Total_Stock = 0
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2



    For i = 2 To Lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker_type = ws.Cells(i, 1).Value
            ws.Range("I" & Summary_Table_Row).Value = Ticker_type


            ' Yearly Change
            Yearly_Change = ws.Cells(i, 6).Value - openingpr
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

            ' Percent Change
            Percent_Change = Yearly_Change / openingpr
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change


            ' Total Stock Volume
            Total_Stock = Total_Stock + ws.Cells(i, 7).Value
            ws.Range("L" & Summary_Table_Row).Value = Total_Stock
            Summary_Table_Row = Summary_Table_Row + 1

            Total_Stock = 0
            openingpr = ws.Cells(i + 1, 3).Value
            
            ' Greatest Value
            ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & Summary_Table_Row)) * 100
            ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & Summary_Table_Row)) * 100
            ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & Summary_Table_Row))
            
            increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & Summary_Table_Row)), ws.Range("K2:K" & Summary_Table_Row), 0)
            decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & Summary_Table_Row)), ws.Range("K2:K" & Summary_Table_Row), 0)
            volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & Summary_Table_Row)), ws.Range("L2:L" & Summary_Table_Row), 0)
            
            ws.Range("P2") = ws.Cells(increase_number + 1, 9)
            ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
            ws.Range("P4") = ws.Cells(volume_number + 1, 9)
            
            
            

        Else

            Total_Stock = Total_Stock + ws.Cells(i, 7).Value



        End If

        ' Yearly Change Colors
        If ws.Cells(i, 10).Value >= 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4

        ElseIf ws.Cells(i, 10).Value < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3

        End If

    Next i



            ws.Columns("K").NumberFormat = "0.00%"
   Next ws
End Sub

