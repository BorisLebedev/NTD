Attribute VB_Name = "DataMKupd"
Sub main()
Dim rng As Range
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("Данные из МК (из pdf)")

With ws
    .Cells.Interior.ColorIndex = 2
    Set rng = .Range(.Cells(DataMK.TOP_INDENT + 1, 1), .Cells(DocumentAttribute.LastRow(ws, 3), DataMK.L_COL))
    Call SameProductsSameColor(rng)
End With


End Sub
