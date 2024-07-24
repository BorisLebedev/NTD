Attribute VB_Name = "Paths"
Function NormAllPath()
NormAllPath = Application.ThisWorkbook.Worksheets("Расчет").Cells(1, 1).Value2
If NormAllPath = "" Then
    wbpath = Application.ThisWorkbook.path
    lenght = InStrRev(wbpath, "\", -1, 0)
    NormAllPath = Left(wbpath, lenght)
End If
End Function

Function NormAllName()
NormAllName = Application.ThisWorkbook.Worksheets("Расчет").Cells(1, 4).Value2
End Function
