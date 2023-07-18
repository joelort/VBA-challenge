Attribute VB_Name = "Module1"

Sub credit_charges()

    Worksheets("2018").Activate

     Cells(1, 10).Value = "Ticker"
     Cells(1, 11).Value = "Year change"
     Cells(1, 12).Value = "Percent change"
     Cells(1, 13).Value = "Total stock volume"

Worksheets("2019").Activate

     Cells(1, 10).Value = "Ticker"
     Cells(1, 11).Value = "Year change"
     Cells(1, 12).Value = "Percent change"
     Cells(1, 13).Value = "Total stock volume"
     
     Worksheets("2020").Activate

     Cells(1, 10).Value = "Ticker"
     Cells(1, 11).Value = "Year change"
     Cells(1, 12).Value = "Percent change"
     Cells(1, 13).Value = "Total stock volume"
End Sub

Sub test2()
Dim row As Long
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
row = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
ws.Range("A2:A" & row).AdvancedFilter _
Action:=xlFilterCopy, CopyToRange:=ws.Range("J2"), _
Unique:=True
Next ws

End Sub

