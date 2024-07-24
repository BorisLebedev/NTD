Attribute VB_Name = "Paths"
Function NormAllPath()
NormAllPath = Application.ThisWorkbook.Worksheets("Настройки").Cells(1, 2).Value2
If NormAllPath = "" Then
    wbpath = Application.ThisWorkbook.path
    lenght = InStrRev(wbpath, "\", -1, 0)
    NormAllPath = Left(wbpath, lenght)
End If
End Function

Function NormAllName()
NormAllName = Application.ThisWorkbook.Worksheets("Настройки").Cells(2, 2).Value2
End Function

Function OperationsPath()
OperationsPath = Application.ThisWorkbook.Worksheets("Настройки").Cells(3, 2).Value2
If OperationsPath = "" Then
    wbpath = Application.ThisWorkbook.path
    lenght = InStrRev(wbpath, "\", -1, 0)
    OperationsPath = Left(wbpath, lenght)
End If
End Function

Function OperationsName()
OperationsName = Application.ThisWorkbook.Worksheets("Настройки").Cells(4, 2).Value2
End Function

