Attribute VB_Name = "Consolidation"
Global Const OPERATION_ERROR_MSG As String = Calculation.OPERATION_ERROR_MSG
Global Const col_level As Integer = Calculation.col_level
Global Const col_hierarchy As Integer = Calculation.col_hierarchy
Global Const col_name As Integer = Calculation.col_name
Global Const col_deno As Integer = Calculation.col_deno
Global Const col_deno_td As Integer = Calculation.col_deno_td
Global Const col_operation As Integer = Calculation.col_operation
Global Const col_num As Integer = Calculation.col_num
Global Const col_norm As Integer = Calculation.col_norm
Global Const col_norm_total As Integer = Calculation.col_norm_total
Global Const col_base As Integer = Calculation.col_base
Global Const col_norm_calc As Integer = Calculation.col_norm_calc
Global Const col_norm_fix As Integer = Calculation.col_norm_fix
Global Const l_col As Integer = Calculation.l_col
Global Const top_indent As Integer = Calculation.top_indent

Global Const col_level_cons As Integer = 1
Global Const col_hierarchy_cons As Integer = col_level_cons + 1
Global Const col_name_cons As Integer = col_hierarchy_cons + 1
Global Const col_norm_cons As Integer = col_name_cons + 1
Global Const col_num_cons As Integer = col_norm_cons + 1
Global Const col_weight_cons As Integer = col_num_cons + 1
Global Const l_col_cons As Integer = col_weight_cons + 1


Sub Main()
Const index_col As Integer = col_name
Const consolidation_ws_name As String = "Консолидация"
Dim data As Variant
Dim ws As Worksheet
Dim l_row As Long
Dim rng As Range
Dim level As Integer
Dim data0 As Variant
Dim data1 As Variant
Dim data2 As Variant

Call Screen.Events(False)
Set ws = ActiveSheet()
Calculation.OPERATIONS = Calculation.GetOperationArray()
Calculation.OPERATIONS_CORRECTION = Calculation.GetOperationCorrectionArray()
Calculation.OPERATIONS_TYPE_ORDER = Calculation.GetOperationTypeOrderArray()
With ws
    l_row = DocumentAttribute.LastRow(ws, col_norm) + 1
    If l_row > top_indent Then
        Set rng = .Range(.Cells(top_indent + 1, 1), .Cells(l_row, l_col))
    Else
        Set rng = .Range(.Cells(top_indent + 1, 1), .Cells(top_indent + 1, l_col))
    End If
    data = rng.value
End With

data = ClearColumn(data, col_deno_td)
data = CalculateWeights(data)
data = CalculateConsolidation(data)


DocumentAttribute.AddNewWS (consolidation_ws_name)
Set ws_consolidation = ThisWorkbook.Worksheets(consolidation_ws_name)
With ws_consolidation
    .Cells.ClearContents
    .Cells.Font.Bold = False
    .Cells(1, col_level_cons).EntireColumn.ColumnWidth = 5
    .Cells(1, col_hierarchy_cons).EntireColumn.NumberFormat = "@"
    .Cells(1, col_hierarchy_cons).EntireColumn.ColumnWidth = 10
    .Cells(1, col_name_cons).EntireColumn.ColumnWidth = 80
    .Cells(1, col_norm_cons).EntireColumn.ColumnWidth = 10
    .Cells(1, col_num_cons).EntireColumn.NumberFormat = "0"
    .Cells(1, col_weight_cons).EntireColumn.NumberFormat = "0"
    Set rng = .Range(.Cells(top_indent + 1, col_level_cons), .Cells(top_indent + UBound(data), l_col_cons))
    rng = data
    .Cells(1, step + col_name_cons) = "Консолидация с весом и количеством"
    .Cells(1, step + col_name_cons).Font.Size = 14
    .Cells(1, step + col_name_cons).Font.Bold = True
    .Cells(1, step + col_num_cons) = "Кол-во"
    .Cells(1, step + col_weight_cons) = "Вес"
    rng.EntireColumn.Hidden = True
End With

Call CreateConsolidationByLevel(data)

Call Screen.Events(True)
End Sub


Private Function ClearColumn(data As Variant, column As Integer)

For row = LBound(data) To UBound(data)
    data(row, column) = Empty
Next row

ClearColumn = data

End Function


Private Sub CreateConsolidationByLevel(data As Variant)
Const consolidation_ws_name As String = "Консолидация"
Dim max_level As Integer
Dim data_cons As Variant
Dim level As Integer
Dim step As Integer
Dim rng As Range

max_level = Calculation.SearchMaxLevel(data)
step = 0

Set ws_consolidation = ThisWorkbook.Worksheets(consolidation_ws_name)
With ws_consolidation
    For level = max_level To 0 Step -1
        data_cons = CalculateConsolidationByLevel(data, level)
        
        step = (level + 1) * l_col_cons
        .Cells(1, step + col_hierarchy_cons).EntireColumn.NumberFormat = "@"
        .Cells(1, step + col_level_cons).EntireColumn.ColumnWidth = 10
        .Cells(1, step + col_level_cons).EntireColumn.HorizontalAlignment = xlHAlignCenter
        .Cells(1, step + col_hierarchy_cons).EntireColumn.ColumnWidth = 10
        .Cells(1, step + col_hierarchy_cons).EntireColumn.HorizontalAlignment = xlHAlignCenter
        .Cells(1, step + col_name_cons).EntireColumn.ColumnWidth = 80
        .Cells(1, step + col_norm_cons).EntireColumn.ColumnWidth = 10
        .Cells(1, step + col_norm_cons).EntireColumn.NumberFormat = "0.00"
        .Cells(1, step + col_num_cons).EntireColumn.Hidden = True
        .Cells(1, step + col_weight_cons).EntireColumn.Hidden = True
        Set rng = .Range(.Cells(top_indent + 1, step + col_level_cons), .Cells(top_indent + UBound(data_cons), step + col_norm_cons))
        rng = data_cons
        
        rng.Borders.LineStyle = XlLineStyle.xlContinuous
        Call MakeProductsBold(rng)
        
        .Cells(1, step + col_name_cons) = "Консолидация по уровню " & level
        .Cells(1, step + col_name_cons).Font.Size = 14
        .Cells(1, step + col_name_cons).Font.Bold = True
        
        .Cells(2, step + col_level_cons) = "Уровень"
        .Cells(2, step + col_hierarchy_cons) = "Индекс"
        .Cells(2, step + col_name_cons) = "Наименование / Вид работ"
        .Cells(2, step + col_norm_cons) = "Тр-ть, н/ч"
        
        .Cells(2, step + col_level_cons).Font.Bold = True
        .Cells(2, step + col_hierarchy_cons).Font.Bold = True
        .Cells(2, step + col_name_cons).Font.Bold = True
        .Cells(2, step + col_norm_cons).Font.Bold = True
        
        .Cells(2, step + col_name_cons).HorizontalAlignment = xlHAlignCenter
        .Cells(2, step + col_norm_cons).HorizontalAlignment = xlHAlignCenter

    Next
End With
End Sub


Private Sub MakeProductsBold(rng As Range)
Dim col_level_cell As Integer

col_level_cell = rng.Columns.Count - l_col_cons

For Each row In rng.Rows
    If row.Cells(1, 1).value <> "" Then
        row.Font.Bold = True
    Else: End If
Next row

End Sub


Private Function GetConsLevel(data As Variant, row As Long, current_level As Integer)

If Not IsEmpty(data(row, col_level_cons)) And data(row, col_level_cons) <> "" Then
    GetConsLevel = data(row, col_level_cons)
Else
    GetConsLevel = current_level
End If

End Function


Function CalculateWeights(data As Variant)
Dim row_sub As Long
Dim level As Integer

For row = UBound(data) To LBound(data) Step -1
    If Not IsEmpty(data(row, Calculation.col_level)) And data(row, col_level_cons) <> "" Then
        num = data(row, Calculation.col_num)
        weight = 1
        base_level = data(row, Calculation.col_level)
        next_level = base_level - 1
    
        For row_sub = row To LBound(data) Step -1
            level = GetConsLevel(data, row_sub, level)
            If level = next_level Then
                num = data(row_sub, Calculation.col_num)
                weight = weight * num
                next_level = next_level - 1
            Else: End If
        Next row_sub
        
        data(row, Calculation.col_deno_td) = weight

    Else: End If
Next row
CalculateWeights = data
End Function


Private Function CalculateConsolidationByLevel(data As Variant, level_cons As Integer)
Dim row As Long
Dim l_row As Long
Dim index_hierarchy As String
Dim index_hierarchy_next As String
Dim num As Integer
Dim level As Integer
Dim data_new As Variant
Dim row_new As Long

ReDim data_new(LBound(data, 2) To UBound(data, 2), _
               LBound(data) To LBound(data))
row_new = 1

For row = LBound(data) To UBound(data)
    level = GetConsLevel(data, row, 0)
    
    If level <= level_cons Then
        ReDim Preserve data_new(LBound(data_new) To UBound(data_new), _
                                LBound(data_new, 2) To UBound(data_new, 2) + 1)
        data_new(col_level_cons, row_new) = data(row, col_level_cons)
        data_new(col_hierarchy_cons, row_new) = data(row, col_hierarchy_cons)
        data_new(col_name_cons, row_new) = data(row, col_name_cons)
        data_new(col_norm_cons, row_new) = data(row, col_norm_cons)
        row_new = row_new + 1
        If level = level_cons Then
            row = CalculateJobsByLevel(data, row, level, data_new, row_new)
        Else: End If
    Else
    
    End If
    
Next
data_new = DocumentAttribute.Transpose2dArray(data_new)
'data_new = DocumentAttribute.Reverse2dArray(data_new)

CalculateConsolidationByLevel = data_new
End Function


Private Function CalculateJobsByLevel(data As Variant, row_base As Long, level_base As Integer, data_new As Variant, row_new As Long)
Dim time As Double
Dim operation_type As String
Dim job_type As String
Dim row As Long
Dim order As Integer
Dim l_row As Long
Dim row_next_level As Long

l_row = UBound(data)
row_next_level = l_row
For order = LBound(Calculation.OPERATIONS_TYPE_ORDER) To UBound(Calculation.OPERATIONS_TYPE_ORDER)
    job_type = Calculation.OPERATIONS_TYPE_ORDER(order, 1)
    operation_type = ""
    time = 0
    row = row_base + 1
    level = level_base + 1
    weight = 1
    num = 1
    
    Do While level > level_base And row <= l_row
        If data(row, col_level_cons) = "" Then
            operation_type = data(row, col_name_cons)
            If job_type = operation_type Then
                time = time + data(row, col_norm_cons)
            Else: End If
        Else: End If
        
        If data(row, col_level_cons) <> "" Then
            level = data(row, col_level_cons)
'            weight = data(row, col_weight_cons)
            num = data(row, col_num_cons)
            If level <= level_base Then
                row_next_level = row - 1
            Else: End If
        Else: End If
        row = row + 1
    Loop
    
    If time > 0 Then
        
        data_new(col_name_cons, row_new) = job_type
        data_new(col_norm_cons, row_new) = time
        ReDim Preserve data_new(LBound(data_new) To UBound(data_new), _
                                LBound(data_new, 2) To UBound(data_new, 2) + 1)
        row_new = row_new + 1

    Else: End If
Next
CalculateJobsByLevel = row_next_level
    
End Function


Private Function CalculateConsolidation(data As Variant)
Dim row As Long
Dim l_row As Long
Dim row_new As Long
Dim row_sub As Integer
Dim index_hierarchy As String
Dim index_hierarchy_next As String
Dim num As Double
Dim norm As Double
'Dim time_operation As Double
Dim time_subproduct As Double
Dim level As Integer
Dim level_next As Integer
Dim data_new As Variant
Dim weight As Double


ReDim data_new(1 To l_col_cons, 1 To 1)

row_new = 1
l_row = UBound(data)
For row = UBound(data) To LBound(data) Step -1
    

    index_hierarchy = data(row, col_hierarchy)
    index_hierarchy_next = ""
    row_sub = 0
    num = data(row, col_num)
    weight = data(row, col_deno_td)
    time_operation = 0
    time_subproduct = 0
    level = 0
    level_next = 0
    norm = 0
    If index_hierarchy <> "" Then
'        If index_hierarchy = "12" Then
'            num1 = "1"
'        Else: End If

        ReDim Preserve data_new(LBound(data_new) To UBound(data_new), _
                                LBound(data_new, 2) To UBound(data_new, 2) + 1)

        If Not IsEmpty(data(row, col_norm_fix)) Then
            'Все в ИЭТ (Зафиксированная трудоемкость. Игнорируются все входящие данные)
            norm = num * CDec(data(row, col_norm_fix))
        Else
            level = Calculation.GetLevel(index_hierarchy)
            If row = UBound(data) Then
                'Все в ИЭТ (Последняя строка расшифовки без операций)
                norm = num * Calculation.GetDataNorm(data, row)
            Else
                index_hierarchy_next = data(row + 1, col_hierarchy)
                If index_hierarchy_next <> "" Then
                    level_next = Calculation.GetLevel(index_hierarchy_next)
                    If level_next > level Then
                        'Все в ИЭТ (Входящие с НЕнулевой суммой но операций нет)
                        time_subproduct = Calculation.SumSubProducts(data, row, l_row, level)
                        norm = num * time_subproduct
                        If time_subproduct = 0 Then
                            'Все в ИЭТ (Входящие с нулевой суммой)
                            norm = num * Calculation.GetDataNorm(data, row)
                        Else: End If
                    Else
                        'Все в ИЭТ (Входящих нет)
                        norm = num * Calculation.GetDataNorm(data, row)
                    End If
                Else
                    'Пооперационная сумма
                    time_subproduct = Calculation.SumSubProducts(data, row, l_row, level)
                    'Все в ИЭТ (Входящие)
                    time_operation = SumOperations(data, row, l_row, level, time_subproduct)
                    norm = num * (time_operation + time_subproduct)
                    
                    Call FormatOperationByType(data, row, l_row, data_new, row_new, level, num, weight, time_subproduct)
                    
                    If time_subproduct = 0 And time_operation = 0 Then
                        'Все в ИЭТ (Пооперационная сумма и сумма входящихх нулевые)
                        norm = num * Calculation.GetDataNorm(data, row)
                    Else: End If
                End If
            End If
        End If
        
        If time_subproduct = 0 And time_operation = 0 Then
            data_new(col_name, row_new) = "Сборка и монтаж изделий электронной техники"
            data_new(col_deno, row_new) = weight * data(row, col_norm_calc)
            ReDim Preserve data_new(LBound(data_new) To UBound(data_new), _
                                    LBound(data_new, 2) To UBound(data_new, 2) + 1)
            row_new = row_new + 1
        Else: End If
        
'        If time_subproduct = 0 And time_operation <> 0 Then
'            data_new(col_name, row_new) = "Сборка и монтаж изделий электронной техники"
'            data_new(col_deno, row_new) = data(row, col_norm_calc) - time_operation
'            ReDim Preserve data_new(LBound(data_new) To UBound(data_new), _
'                                    LBound(data_new, 2) To UBound(data_new, 2) + 1)
'            row_new = row_new + 1
'        Else: End If
        
        data_new(col_level_cons, row_new) = data(row, col_level)
        data_new(col_hierarchy_cons, row_new) = data(row, col_hierarchy)
        data_new(col_name_cons, row_new) = data(row, col_name) & " " & data(row, col_deno) & ", " & num * weight & " " & Complectov(num * weight)
        data_new(col_norm_cons, row_new) = weight * data(row, col_norm_calc)
        data_new(col_num_cons, row_new) = num
        data_new(col_weight_cons, row_new) = weight
        
        row_new = row_new + 1
    Else: End If
Next
data_new = DocumentAttribute.Transpose2dArray(data_new)
data_new = DocumentAttribute.Reverse2dArray(data_new)


CalculateConsolidation = data_new
End Function


Private Sub FormatOperationByType(data As Variant, base_row As Long, l_row As Long, data_new As Variant, row_new As Long, level As Integer, num As Double, weight As Double, time_subproduct As Double)

Dim time As Double
Dim operation_name As String
Dim operation_time As Double
Dim operation_type As String
Dim job_type As String
Dim row As Long
Dim order As Integer

For order = UBound(Calculation.OPERATIONS_TYPE_ORDER) To LBound(Calculation.OPERATIONS_TYPE_ORDER) Step -1
    job_type = Calculation.OPERATIONS_TYPE_ORDER(order, 1)
    operation_type = ""
    time = 0
    row = base_row + 1
    
    If base_row = 208 Then
        row = row
    Else: End If
    
'    If Not (job_type = "Сборка и монтаж изделий электронной техники" And time_subproduct <> 0) Then
    operation_name = data(row, col_operation)
    
    Do While data(row, col_hierarchy) = "" And row <> l_row
        If Not (job_type = "Сборка и монтаж изделий электронной техники" And time_subproduct <> 0 And operation_name <> "Сумма операций") Then
            operation_type = data(row, col_deno)
            If job_type = operation_type Then
                product_time = Calculation.GetDataNorm(data, row)
                time = time + product_time
            Else: End If
        Else: End If
        row = row + 1
        operation_name = data(row, col_operation)
    Loop
    If time > 0 Then
        data_new(col_name_cons, row_new) = job_type
        data_new(col_norm_cons, row_new) = weight * num * time
        ReDim Preserve data_new(LBound(data_new) To UBound(data_new), _
                                LBound(data_new, 2) To UBound(data_new, 2) + 1)
        row_new = row_new + 1
    Else: End If
Next

End Sub


Private Function Complectov(num As Long)
Dim word As String

Select Case num

Case 1
    word = "комплект"
Case 2 To 4
    word = "комплекта"
Case 5 To 20
    word = "комплектов"
Case Else
    word = "комплект"
End Select
Complectov = word

End Function
