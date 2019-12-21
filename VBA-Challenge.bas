Sub Stock()
        'Create Variable to hold Value
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearChange As Double
        Dim Ticker As String
        Dim PercentChange As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim SumTable As Integer
        SumTable = 1
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row

        ' Add Heading for summary
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"

        'Set Initial Open Price
        OpenPrice = Cells(2, SumTable + 2).Value
        
		' Loop through all ticker symbol   
        For i = 2 To LastRow
			' Check if we are still within the same ticker (Part 1)
            If Cells(i + 1, SumTable).Value <> Cells(i, SumTable).Value Then
                ' Set Ticker name
                Ticker = Cells(i, SumTable).Value
                Cells(Row, SumTable + 8).Value = Ticker
                ' Set Close Price
                ClosePrice = Cells(i, SumTable + 5).Value
                ' Add Yearly Change
                YearChange = ClosePrice - OpenPrice
                Cells(Row, SumTable + 9).Value = YearChange
                    If (Cells(Row, SumTable + 9).Value < 0) Then
                        Cells(Row, SumTable + 9).Interior.ColorIndex = 3
                    Else
                        Cells(Row, SumTable + 9).Interior.ColorIndex = 4
                    End If
                ' Add Percent Change
                If (OpenPrice = 0 And ClosePrice = 0) Then
                    PercentChange = 0
                ElseIf (OpenPrice = 0 And ClosePrice <> 0) Then
                    PercentChange = 1
                Else
                    PercentChange = YearChange / OpenPrice
                    Cells(Row, SumTable + 10).Value = PercentChange
                    Cells(Row, SumTable + 10).NumberFormat = "0.00%"
                End If
                ' Add Total Volume
                Volume = Volume + Cells(i, SumTable + 6).Value
                Cells(Row, SumTable + 11).Value = Volume
                ' Add one to the summary table row
                Row = Row + 1
                ' reset the Open Price
                OpenPrice = Cells(i + 1, SumTable + 2)
                ' reset the Volumn Total
                Volume = 0
            'If cells are the same ticker (Part 2)
            Else
                Volume = Volume + Cells(i, SumTable + 6).Value
            End If
        Next i
End Sub
''''''''''''''''''''Challenge for looping''''''''''''''''''''''''''''''''
Sub StockChallenge()
        'Loop through all worksheets
        Dim WS As Worksheet
        For Each WS In ActiveWorkbook.Worksheets
        WS.Activate

        'Create Variable to hold Value
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearChange As Double
        Dim Ticker As String
        Dim PercentChange As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim SumTable As Integer
        SumTable = 1
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        ' Add Heading for summary
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"

        'Set Initial Open Price
        OpenPrice = Cells(2, SumTable + 2).Value
        
		' Loop through all ticker symbol
        For i = 2 To LastRow
        
		' Check if we are still within the same ticker (Part 1)
            If Cells(i + 1, SumTable).Value <> Cells(i, SumTable).Value Then
                ' Set Ticker name
                Ticker = Cells(i, SumTable).Value
                Cells(Row, SumTable + 8).Value = Ticker
                ' Set Close Price
                ClosePrice = Cells(i, SumTable + 5).Value
                ' Add Yearly Change
                YearChange = ClosePrice - OpenPrice
                Cells(Row, SumTable + 9).Value = YearChange
                    If (Cells(Row, SumTable + 9).Value < 0) Then
                        Cells(Row, SumTable + 9).Interior.ColorIndex = 3
                    Else
                        Cells(Row, SumTable + 9).Interior.ColorIndex = 4
                    End If
                ' Add Percent Change
                If (OpenPrice = 0 And ClosePrice = 0) Then
                    PercentChange = 0
                ElseIf (OpenPrice = 0 And ClosePrice <> 0) Then
                    PercentChange = 1
                Else
                    PercentChange = YearChange / OpenPrice
                    Cells(Row, SumTable + 10).Value = PercentChange
                    Cells(Row, SumTable + 10).NumberFormat = "0.00%"
                End If
                ' Add Total Volume
                Volume = Volume + Cells(i, SumTable + 6).Value
                Cells(Row, SumTable + 11).Value = Volume
                ' Add one to the summary table row
                Row = Row + 1
                ' reset the Open Price
                OpenPrice = Cells(i + 1, SumTable + 2)
                ' reset the Volumn Total
                Volume = 0
            'If cells are the same ticker (Part 2)
            Else
                Volume = Volume + Cells(i, SumTable + 6).Value
            End If
        Next i
    Next WS
End Sub
