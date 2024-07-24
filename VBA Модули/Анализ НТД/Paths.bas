Attribute VB_Name = "Paths"
Function NormAllPath()
NormAllPath = Application.ThisWorkbook.Worksheets("Настройки").Cells(2, 2).Value2
If NormAllPath = "" Then
    NormAllPath = Application.ThisWorkbook.Path & "\"
End If
End Function

Function NormAllName()
NormAllName = Application.ThisWorkbook.Worksheets("Настройки").Cells(1, 2).Value2
End Function

Function NTDPath()
NTDPath = Application.ThisWorkbook.Worksheets("Настройки").Cells(3, 2).Value2
If NTDPath = "" Then
    NTDPath = Application.ThisWorkbook.Path & "\Данные о трудоемкости ремонта"
End If
End Function


