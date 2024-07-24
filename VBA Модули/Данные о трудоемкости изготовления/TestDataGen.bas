Attribute VB_Name = "TestDataGen"
Sub Main()
Const index_col As Integer = 7
Dim data As Variant
Dim new_data As Variant
Dim ws As Worksheet
Dim l_row As Long
Dim rng As Range
Dim rng_base As Range
Dim operation_arr As Variant

Call Screen.Events(False)
Set ws = ActiveSheet()
With ws
    
    l_row = DocumentAttribute.LastRow(ws, Calculation.col_name)
    If l_row > Calculation.top_indent Then
        Set rng = .Range(.Cells(Calculation.top_indent + 1, 1), .Cells(l_row, Calculation.l_col))
    Else
        Set rng = .Range(.Cells(Calculation.top_indent + 1, 1), .Cells(Calculation.top_indent + 1, Calculation.l_col))
    End If
    data = rng.value
    ReDim new_data(LBound(data, 2) To UBound(data, 2), 1 To 1)
    
    For row = LBound(data) To UBound(data)
        operation_set = Int(3 * Rnd + 1)
    
        Select Case operation_set
        Case 1
            ReDim operation_arr(1 To 6)
            operation_arr(1) = "Комплектование"
            operation_arr(2) = "Подготовка"
            operation_arr(3) = "Электромонтаж"
            operation_arr(4) = "Сборка"
            operation_arr(5) = "Контроль"
            operation_arr(6) = "Складирование"
        Case 2
            ReDim operation_arr(1 To 4)
            operation_arr(1) = "Комплектование"
            operation_arr(2) = "Сборка"
            operation_arr(3) = "Контроль"
            operation_arr(4) = "Складирование"
        Case 3
            ReDim operation_arr(1 To 5)
            operation_arr(1) = "Комплектование"
            operation_arr(2) = "Подготовка"
            operation_arr(3) = "Сборка"
            operation_arr(4) = "Контроль"
            operation_arr(5) = "Складирование"
        Case 4
            ReDim operation_arr(1 To 5)
            operation_arr(1) = "Комплектование"
            operation_arr(2) = "Подготовка"
            operation_arr(3) = "Электромонтаж"
            operation_arr(4) = "Контроль"
            operation_arr(5) = "Складирование"
        End Select
        
        
        new_row = UBound(new_data, 2)
        For col = LBound(new_data) To UBound(new_data)
            new_data(col, new_row) = data(row, col)
        Next col
        
        ReDim Preserve new_data(LBound(new_data) To UBound(new_data), _
                                LBound(new_data, 2) To UBound(new_data, 2) + UBound(operation_arr))
        For sub_row = LBound(operation_arr) To UBound(operation_arr)
            new_data(1, new_row + sub_row) = ""
            new_data(2, new_row + sub_row) = ""
            new_data(3, new_row + sub_row) = operation_arr(sub_row)
            new_data(4, new_row + sub_row) = ""
            new_data(5, new_row + sub_row) = ""
            new_data(6, new_row + sub_row) = ""
            new_data(7, new_row + sub_row) = CDbl(format(Rnd, "#,##0.00"))
            new_data(8, new_row + sub_row) = ""
            new_data(9, new_row + sub_row) = ""
            new_data(10, new_row + sub_row) = ""
            new_data(11, new_row + sub_row) = ""
            new_data(12, new_row + sub_row) = ""
        Next sub_row
        ReDim Preserve new_data(LBound(new_data) To UBound(new_data), LBound(new_data, 2) To UBound(new_data, 2) + 1)
    Next row
    
    new_data = DocumentAttribute.Transpose2dArray(new_data)
    Set rng = .Range(.Cells(Calculation.top_indent + 1, 1), .Cells(UBound(new_data), Calculation.l_col))
    rng = new_data
End With
    
Call Screen.Events(True)
End Sub
