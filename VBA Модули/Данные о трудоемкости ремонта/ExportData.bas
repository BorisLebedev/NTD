Attribute VB_Name = "ExportData"
Global Const col_hierarchy As Integer = 1

Global Const col_name As Integer = col_hierarchy + 1
Global Const col_deno As Integer = col_name + 1
Global Const col_num As Integer = col_deno + 1
Global Const col_msr As Integer = col_num + 1

Global Const col_def_one As Integer = col_msr + 1
Global Const col_dis_one As Integer = col_def_one + 1
Global Const col_ass_one As Integer = col_dis_one + 1
Global Const col_rpr_one As Integer = col_ass_one + 1
Global Const col_rpl_one As Integer = col_rpr_one + 1
Global Const col_tun_one As Integer = col_rpl_one + 1
Global Const col_new_one As Integer = col_tun_one + 1

Global Const col_type As Integer = col_new_one + 1

Global Const l_col As Integer = col_new_one
Global Const top_indent As Integer = 12

Sub Main()
Const index_col As Integer = col_name
Const l_row_string As String = "* Ремонт не возможен"
Dim values_name As Variant
Dim values_data As Variant
Dim values_type As Variant
Dim ws As Worksheet
Dim ws_ntd As Worksheet
Dim l_row As Long
Dim l_row_ntd As Long
Dim rng_name As Range
Dim rng_name_ntd As Range
Dim rng_data As Range
Dim rng_data_ntd As Range
Dim rng_type As Range
Dim rng_type_ntd As Range

Dim currentFiltRange As String
Dim filterArray As Variant

Call Screen.Events(False)
Set ws = ThisWorkbook.Worksheets("Расчет")
With ws
    Call Screen.SaveAutoFilter(ws, currentFiltRange, filterArray)
    .AutoFilter.ShowAllData
    l_row = DocumentAttribute.LastRow(ws, col_name)
    If l_row > Calculation.top_indent Then
        .Cells(Calculation.top_indent + 1, Calculation.col_hierarchy) = "Изделие"
        Set rng_name = .Range(.Cells(Calculation.top_indent + 1, Calculation.col_hierarchy), .Cells(l_row, Calculation.col_deno))
        Set rng_data = .Range(.Cells(Calculation.top_indent + 1, Calculation.col_num), .Cells(l_row, Calculation.col_new_one))
        Set rng_type = .Range(.Cells(Calculation.top_indent + 1, Calculation.col_type), .Cells(l_row, Calculation.col_type))
    Else
        Set rng_name = .Range(.Cells(Calculation.top_indent + 1, Calculation.col_hierarchy), .Cells(Calculation.top_indent + 1, Calculation.col_deno))
        Set rng_data = .Range(.Cells(Calculation.top_indent + 1, Calculation.col_num), .Cells(Calculation.top_indent + 1, Calculation.col_new_one))
        Set rng_type = .Range(.Cells(Calculation.top_indent + 1, Calculation.col_type), .Cells(Calculation.top_indent + 1, Calculation.col_type))
    End If
    values_name = rng_name.value
    values_data = rng_data.value
    values_type = rng_type.value
    Call Screen.RestoreAutoFilter(ActiveSheet, currentFiltRange, filterArray)
End With

Set ws_ntd = ThisWorkbook.Worksheets("НТД")
With ws_ntd
    l_row_ntd = .Range("A:AA").Find(l_row_string, .Cells(1, 1)).row - 3
        
    If l_row_ntd > top_indent Then
        .Range(.Cells(top_indent + 1, 1), .Cells(l_row_ntd, 1)).EntireRow.Delete
    End If
    
    .Range(.Cells(top_indent + 1, 1), .Cells(top_indent + UBound(values_data), col_new_one - col_hierarchy)).EntireRow.Insert
    
    Set rng_name_ntd = .Range(.Cells(top_indent + 1, 1), .Cells(top_indent + UBound(values_name), col_deno))
    Set rng_data_ntd = .Range(.Cells(top_indent + 1, col_num), .Cells(top_indent + UBound(values_data), col_new_one))
    Set rng_type_ntd = .Range(.Cells(top_indent + 1, col_type), .Cells(top_indent + UBound(values_data), col_type))
    
    rng_name_ntd.value = values_name
    rng_data_ntd.value = values_data
    rng_type_ntd.value = values_type
    
    rng_name_ntd.RowHeight = 31.5
    
    rng_name_ntd.HorizontalAlignment = xlLeft
    rng_name_ntd.VerticalAlignment = xlCenter
    
    rng_data_ntd.HorizontalAlignment = xlCenter
    rng_data_ntd.VerticalAlignment = xlCenter
    
    rng_type_ntd.HorizontalAlignment = xlCenter
    rng_type_ntd.VerticalAlignment = xlCenter
    
    rng_name_ntd.Borders(xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
    rng_name_ntd.Borders(xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
    
    rng_data_ntd.Borders(xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
    rng_data_ntd.Borders(xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
    
    rng_name_ntd.Borders(xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
    rng_name_ntd.Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
    rng_name_ntd.Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
    rng_name_ntd.Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
    
    rng_data_ntd.Borders(xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
    rng_data_ntd.Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
    rng_data_ntd.Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
    rng_data_ntd.Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
End With

Call Screen.Events(True)
End Sub
