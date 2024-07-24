Attribute VB_Name = "TypeOfProduct"
Global Const col_name As Integer = 2

Global Const col_def As Integer = 3
Global Const col_dis As Integer = 4
Global Const col_ass As Integer = 5
Global Const col_rpr As Integer = 6
Global Const col_rpl As Integer = 7
Global Const col_tun As Integer = 8
Global Const col_new As Integer = 9

Global Const l_col As Integer = col_new

Function GetBaseValues(data As Variant, row_data As Long)
Dim ws As Worksheet


Set ws = ThisWorkbook.Worksheets("Типы")
l_row = DocumentAttribute.LastRow(ws, 1)

With ws
    Set rng = .Range(.Cells(3, 1), .Cells(l_row, l_col))
    rules_data = rng.value
End With

For row = LBound(rules_data) To UBound(rules_data)
    If rules_data(row, col_name) = data(row_data, Calculation.col_type) Then
        data(row_data, Calculation.col_def_one_calc) = rules_data(row, col_def)
        data(row_data, Calculation.col_dis_one_calc) = rules_data(row, col_dis)
        data(row_data, Calculation.col_ass_one_calc) = rules_data(row, col_ass)
        data(row_data, Calculation.col_rpr_one_calc) = rules_data(row, col_rpr)
        data(row_data, Calculation.col_rpl_one_calc) = rules_data(row, col_rpl)
        data(row_data, Calculation.col_tun_one_calc) = rules_data(row, col_tun)
        data(row_data, Calculation.col_new_one_calc) = rules_data(row, col_new)
        Exit For
    End If
Next row
GetBaseValues = data
End Function

Function SetBaseValue(row_calc As Long)
Dim ws As Worksheet

Set ws = ThisWorkbook.Worksheets("Типы")
l_row = DocumentAttribute.LastRow(ws, 1)

With ws
    Set rng = .Range(.Cells(3, 1), .Cells(l_row, l_col))
    rules_data = rng.value
End With

Set ws_calc = ThisWorkbook.Worksheets("Расчет")

With ws_calc
    .Cells(row_calc, Calculation.col_def_one_calc) = ""
    .Cells(row_calc, Calculation.col_dis_one_calc) = ""
    .Cells(row_calc, Calculation.col_ass_one_calc) = ""
    .Cells(row_calc, Calculation.col_rpr_one_calc) = ""
    .Cells(row_calc, Calculation.col_rpl_one_calc) = ""
    .Cells(row_calc, Calculation.col_tun_one_calc) = ""
    .Cells(row_calc, Calculation.col_new_one_calc) = ""
    
    For row = LBound(rules_data) To UBound(rules_data)
            
        If rules_data(row, col_name) = .Cells(row_calc, Calculation.col_type) Then
            .Cells(row_calc, Calculation.col_def_one_calc) = rules_data(row, col_def)
            .Cells(row_calc, Calculation.col_dis_one_calc) = rules_data(row, col_dis)
            .Cells(row_calc, Calculation.col_ass_one_calc) = rules_data(row, col_ass)
            .Cells(row_calc, Calculation.col_rpr_one_calc) = rules_data(row, col_rpr)
            .Cells(row_calc, Calculation.col_rpl_one_calc) = rules_data(row, col_rpl)
            .Cells(row_calc, Calculation.col_tun_one_calc) = rules_data(row, col_tun)
            .Cells(row_calc, Calculation.col_new_one_calc) = rules_data(row, col_new)
            Exit For
        End If
    Next row
End With

End Function
