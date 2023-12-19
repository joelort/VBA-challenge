Attribute VB_Name = "Module1"
Sub MultipleYear()
    Dim ws As Worksheet
    For Each ws In Worksheets
        Dim WorksheetName As String
        Dim i As Long
        Dim j As Long
        Dim TickCount As Long
        Dim LastRowA As Long
        Dim LastRowI As Long
        Dim PerChange As Double
        Dim GreatIncr As Double
        Dim GreatDecr As Double
        Dim GreatVol As Double

        WorksheetName = ws.Name

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volunme"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Values"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

        TickCount = 2
        j = 2

        LastRowA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If LastRowA < 2 Then Exit Sub

        For i = 2 To LastRowA
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value

                If ws.Cells(TickCount, 10).Value < 0 Then
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 4
                End If

                If ws.Cells(j, 3).Value <> 0 Then
                    PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    ws.Cells(TickCount, 11).Value = Format(PerChange, "Percent")
                Else
                    ws.Cells(TickCount, 11).Value = Format(0, "Percent")
                End If

                ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))

                TickCount = TickCount + 1
                j = i + 1
            End If
        Next i

        LastRowI = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row

        GreatVol = ws.Cells(2, 12).Value
        GreatIncr = ws.Cells(2, 11).Value
        GreatDecr = ws.Cells(2, 11).Value

        For i = 2 To LastRowI + 1
            If ws.Cells(i, 12).Value > GreatVol Then
                GreatVol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            End If

            If ws.Cells(i, 11).Value > GreatIncr Then
                GreatIncr = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            End If

            If ws.Cells(i, 11).Value < GreatDecr Then
                GreatDecr = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            End If
        Next i

        ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
        ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
        ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
    Next ws
End Sub


