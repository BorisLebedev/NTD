Attribute VB_Name = "Calculation"
Global Const OPERATION_ERROR_MSG As String = "ОПЕРАЦИЯ НЕ НАЙДЕНА"
Global OPERATIONS As Variant
Global OPERATIONS_CORRECTION As Variant
Global OPERATIONS_TYPE_ORDER As Variant
Global Const col_level As Integer = 1
Global Const col_hierarchy As Integer = col_level + 1
Global Const COL_NAME As Integer = col_hierarchy + 1
Global Const COL_DENO As Integer = COL_NAME + 1
Global Const col_deno_td As Integer = COL_DENO + 1
Global Const col_operation As Integer = COL_NAME
Global Const col_num As Integer = col_deno_td + 1
Global Const COL_NORM As Integer = col_num + 1
Global Const col_norm_total As Integer = COL_NORM + 1
Global Const col_base As Integer = col_norm_total + 1
Global Const COL_NORM_CALC As Integer = col_base + 1
Global Const col_norm_fix As Integer = COL_NORM_CALC + 1
Global Const L_COL As Integer = col_norm_fix + 1
Global Const TOP_INDENT As Integer = 1

Sub main()
Const index_col As Integer = COL_NAME
Dim data As Variant
Dim ws As Worksheet
Dim l_row As Long
Dim rng As Range
Dim rng_calc As Range
Dim rng_jobs As Range

Call Screen.Events(False)
OPERATIONS = GetOperationArray()
OPERATIONS_CORRECTION = GetOperationCorrectionArray()
Set ws = ActiveSheet()
With ws
    l_row = DocumentAttribute.LastRow(ws, COL_NAME) + 1
    If l_row > TOP_INDENT Then
        Set rng = .Range(.Cells(TOP_INDENT + 1, 1), .Cells(l_row, L_COL))
    Else
        Set rng = .Range(.Cells(TOP_INDENT + 1, 1), .Cells(TOP_INDENT + 1, L_COL))
    End If
    data = rng.value
    data = RepearData(data)
    data = NormCulc(data)
    
    data_jobs = GetSubArray(data, COL_DENO)
    Set rng_jobs = .Range(.Cells(TOP_INDENT + 1, COL_DENO), .Cells(l_row, COL_DENO))
    rng_jobs = data_jobs
    
    data_calc = GetSubArray(data, COL_NORM_CALC)
    Set rng_calc = .Range(.Cells(TOP_INDENT + 1, COL_NORM_CALC), .Cells(l_row, COL_NORM_CALC))
    rng_calc = data_calc
'    rng = data
    Call SetFormatOfData(rng, data)
End With
Call Screen.Events(True)
End Sub


Function GetSubArray(data As Variant, column As Integer)
Dim data_calc As Variant

ReDim data_calc(LBound(data) To UBound(data), 1 To 1)

For row = LBound(data) To UBound(data)
    data_calc(row, 1) = data(row, column)
Next row
GetSubArray = data_calc

End Function


Function GetOperationArray()
Dim ws As Worksheet
Set ws = ThisWorkbook().Worksheets("Операции")
With ws
    l_row = DocumentAttribute.LastRow(ws, 1) + 1
    If l_row > TOP_INDENT Then
        Set rng = .Range(.Cells(1, 1), .Cells(l_row, 2))
    Else
        Set rng = .Range(.Cells(1, 1), .Cells(1, 2))
    End If
    GetOperationArray = rng.value
End With
End Function


Function GetOperationCorrectionArray()
Dim ws As Worksheet
Set ws = ThisWorkbook().Worksheets("Исправления")
With ws
    l_row = DocumentAttribute.LastRow(ws, 1) + 1
    If l_row > TOP_INDENT Then
        Set rng = .Range(.Cells(1, 1), .Cells(l_row, 2))
    Else
        Set rng = .Range(.Cells(1, 1), .Cells(1, 2))
    End If
    GetOperationCorrectionArray = rng.value
End With
End Function


Function GetOperationTypeOrderArray()
Dim ws As Worksheet
Set ws = ThisWorkbook().Worksheets("Порядок видов работ")
With ws
    l_row = DocumentAttribute.LastRow(ws, 1) + 1
    If l_row > TOP_INDENT Then
        Set rng = .Range(.Cells(1, 1), .Cells(l_row, 1))
    Else
        Set rng = .Range(.Cells(1, 1), .Cells(1, 1))
    End If
    GetOperationTypeOrderArray = rng.value
End With
End Function


Private Function NormCulc(data As Variant)
Dim row As Long
Dim l_row As Long
Dim row_sub As Integer
Dim index_hierarchy As String
Dim index_hierarchy_next As String
Dim num As Integer
Dim norm As Double
Dim time_operation As Double
Dim time_subproduct As Double
Dim level As Integer
Dim level_next As Integer

l_row = UBound(data)
For row = UBound(data) To LBound(data) Step -1
    index_hierarchy = data(row, col_hierarchy)
    index_hierarchy_next = ""
    row_sub = 0
    num = data(row, col_num)
    time_operation = 0
    time_subproduct = 0
    level = 0
    level_next = 0
    norm = 0
    If index_hierarchy <> "" Then
'        If index_hierarchy = "Изделие" Then
'            num1 = "1"
'        Else: End If

        If Not IsEmpty(data(row, col_norm_fix)) Then
            norm = num * CDec(data(row, col_norm_fix))
        Else
            level = GetLevel(index_hierarchy)
            If row = UBound(data) Then
                norm = num * GetDataNorm(data, row)
            Else
                index_hierarchy_next = data(row + 1, col_hierarchy)
                If index_hierarchy_next <> "" Then
                    level_next = GetLevel(index_hierarchy_next)
                    If level_next > level Then
                        time_subproduct = SumSubProducts(data, row, l_row, level)
                        norm = num * time_subproduct
                        If time_subproduct = 0 Then
                            norm = num * GetDataNorm(data, row)
                        Else: End If
                    Else
                        norm = num * GetDataNorm(data, row)
                    End If
                Else
                    time_subproduct = SumSubProducts(data, row, l_row, level)
                    time_operation = SumOperations(data, row, l_row, level, time_subproduct)
                    norm = num * (time_operation + time_subproduct)
                    If time_subproduct = 0 And time_operation = 0 Then
                        norm = num * GetDataNorm(data, row)
                    Else: End If
                End If
            End If
        End If
        data(row, COL_NORM_CALC) = norm
    Else: End If
Next

NormCulc = data
End Function


Function SumSubProducts(data As Variant, base_row As Long, l_row As Long, level As Integer)

Dim time As Double
Dim product_name As String
Dim product_time As Double
Dim index_hierarchy As String
Dim row As Long

time = 0
row = base_row + 1
index_hierarchy = data(row, col_hierarchy)
sub_level = GetLevel(index_hierarchy)
Do While (sub_level > level Or IsEmpty(sub_level)) And row <= l_row
    If Not IsEmpty(sub_level) And sub_level = level + 1 Then
        product_name = data(row, COL_NAME)
        product_time = GetCalcNorm(data, row)
'        product_num = data(row, col_num)
        time = time + product_time
    Else: End If
    row = row + 1
    If row <= l_row Then
        index_hierarchy = data(row, col_hierarchy)
        sub_level = GetLevel(index_hierarchy)
    Else: End If
Loop
SumSubProducts = time

End Function


Function SumOperations(data As Variant, base_row As Long, l_row As Long, level As Integer, time_subproduct As Double)

Dim time As Double
Dim operation_name As String
Dim operation_time As Double
Dim job_type As String
Dim row As Long

time = 0
row = base_row + 1
Do While data(row, col_hierarchy) = "" And row <> l_row
    operation_name = data(row, col_operation)
    operation_time = GetDataNorm(data, row)
    job_type = GetJobType(operation_name, data, row)
     
    If Not (job_type = "Сборка и монтаж изделий электронной техники" And time_subproduct <> 0 And operation_name <> "Сумма операций") Then
        time = time + operation_time
    Else: End If
    
    row = row + 1
Loop
SumOperations = time

End Function


Function GetJobType(operation_name As String, data As Variant, row As Long)
Dim job_type As String

job_type = data(row, COL_DENO)
If row = 180 Then
    row = row
Else: End If

'If (IsEmpty(data(row, col_deno)) Or data(row, col_deno) = OPERATION_ERROR_MSG) And operation_name <> "" Then
If (IsEmpty(data(row, col_hierarchy)) _
Or data(row, COL_DENO) = OPERATION_ERROR_MSG _
Or data(row, col_hierarchy) = "") And operation_name <> "" Then
    job_type = GetJob(operation_name)
Else
    job_type = data(row, COL_DENO)
End If
data(row, COL_DENO) = job_type
GetJobType = job_type

End Function


Function GetJob(operation_name As String)
Dim job_type As String

job_type = OPERATION_ERROR_MSG

For row = LBound(OPERATIONS) To UBound(OPERATIONS)
    If operation_name = OPERATIONS(row, 1) Then
        job_type = OPERATIONS(row, 2)
        Exit For
    Else: End If
Next row

If job_type = OPERATION_ERROR_MSG Then
    operation_name = GetCorrectOperation(operation_name)
    For row = LBound(OPERATIONS) To UBound(OPERATIONS)
        If operation_name = OPERATIONS(row, 1) Then
            job_type = OPERATIONS(row, 2)
            Exit For
        Else: End If
    Next row
Else: End If

GetJob = job_type
End Function

Function GetCorrectOperation(operation_name As String)
Dim new_operation_name As String

new_operation_name = OPERATION_ERROR_MSG

For row = LBound(OPERATIONS_CORRECTION) To UBound(OPERATIONS_CORRECTION)
    If operation_name = OPERATIONS_CORRECTION(row, 1) Then
        new_operation_name = OPERATIONS_CORRECTION(row, 2)
        Exit For
    Else: End If
Next row
GetCorrectOperation = new_operation_name
End Function


Function GetLevel(index_hierarchy As String)
Dim level As Integer

Select Case index_hierarchy
Case "Изделие"
    GetLevel = 0
Case Is <> ""
    If Right(index_hierarchy, 1) = "." Then
        level = Len(index_hierarchy) - Len(Replace(index_hierarchy, ".", ""))
    Else
        level = Len(index_hierarchy) - Len(Replace(index_hierarchy, ".", "")) + 1
    End If
    GetLevel = level
Case Else
    GetLevel = Empty
End Select

End Function


Function GetDataNorm(data As Variant, row As Long)
Dim time As Double
If IsNumeric(data(row, COL_NORM)) Then
    time = data(row, COL_NORM)
Else
    time = 0
End If
GetDataNorm = time
End Function


Function GetLevelNext(data As Variant, base_row As Long, l_row As Long, level As Integer)
Dim row As Long
Dim index_hierarchy As String

row = base_row
Do While IsEmpty(level_next) And row <= l_row - 1
    row = row + 1
    index_hierarchy = data(row, col_hierarchy)
    level_next = GetLevel(index_hierarchy)
Loop
GetLevelNext = level_next

End Function


Function GetCalcNorm(data As Variant, row As Long)
Dim time As Double
If IsNumeric(data(row, COL_NORM_CALC)) Then
    time = data(row, COL_NORM_CALC)
Else
    time = 0
End If
GetCalcNorm = time
End Function


Private Sub SetFormatOfData(rng As Range, data As Variant)

Call RowHeights(rng, data)

For row = UBound(data) To LBound(data) Step -1
    norm = CDec(data(row, col_norm_total))
    norm_calc = CDec(data(row, COL_NORM_CALC))
    
    rng.Cells(row, col_norm_total).Interior.color = RGB(255, 255, 255)
    If Not IsEmpty(norm) Or Not IsEmpty(norm_calc) Then
        If norm <> norm_calc And data(row, col_hierarchy) <> "" Then
            rng.Cells(row, col_norm_total).Interior.color = RGB(255, 0, 0)
        Else: End If
    Else: End If
Next row
End Sub

Function FormatIsError(data As Variant, row As Long, col As Long, format As String)



End Function


Sub RowHeights(rng As Range, data As Variant)

rng.RowHeight = 15
For row = UBound(data) To LBound(data) Step -1
    If data(row, col_hierarchy) <> "" Then
        rng.Cells(row, col_norm_total).RowHeight = 30
    Else: End If
Next row

End Sub

Private Function RepearData(data As Variant)

For row = UBound(data) To 2 Step -1
    data(row, col_hierarchy) = Replace(data(row, col_hierarchy), ",", ".")
Next row
RepearData = data

End Function


Function SearchMaxLevel(data As Variant)
Dim level As Integer
level = 0
For row = LBound(data) To UBound(data)
    If data(row, col_level) > level Then
        level = data(row, col_level)
    Else: End If
Next
SearchMaxLevel = level
End Function
