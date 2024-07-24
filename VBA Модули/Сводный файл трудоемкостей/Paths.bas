Attribute VB_Name = "Paths"
Function CalculationPath()
CalculationPath = Application.ThisWorkbook.Worksheets("Настройки").Cells(1, 2).Value2
If CalculationPath = "" Then
    CalculationPath = Application.ThisWorkbook.Path & "\Данные о трудоемкости изготовления\"
End If
End Function

Function DocMkPath()
DocMkPath = Application.ThisWorkbook.Worksheets("Настройки").Cells(2, 2).Value2
If DocMkPath = "" Then
    DocMkPath = Application.ThisWorkbook.Path & "\Маршрутные карты\"
End If
End Function

Function OperationsPath()
OperationsPath = Application.ThisWorkbook.Worksheets("Настройки").Cells(3, 2).Value2
If OperationsPath = "" Then
    wbpath = Application.ThisWorkbook.Path
    lenght = InStr(1, wbpath, "NTD", 0)
    OperationsPath = Application.ThisWorkbook.Path & "\"
End If
End Function

Function OperationsName()
OperationsName = Application.ThisWorkbook.Worksheets("Настройки").Cells(4, 2).Value2
End Function

