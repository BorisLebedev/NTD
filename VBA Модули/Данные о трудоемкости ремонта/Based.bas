Attribute VB_Name = "Based"

Private Function GetDenoFromCulc(data As Variant, row As Long)
Dim text As String
Dim deno As String

text = data(row, Calculation.col_deno)
deno = WsSubs.FindDeno(text, True)

If deno = "" Then
    GetDenoFromCulc = text
Else
    GetDenoFromCulc = deno
End If

End Function


Private Function ExistInArray(arr As Variant, val As String)
Dim exist As Boolean

exist = False
For Each item In arr
    If item = val Then
        exist = True
    Else: End If
Next item

ExistInArray = exist

End Function


Private Function AddToArray(arr As Variant, val As String)

ReDim Preserve arr(1 To UBound(arr) + 1)
arr(UBound(arr)) = val
AddToArray = arr

End Function


Private Function ArrayToStr(arr As Variant)
Dim text As String

text = ""
For Each item In arr
    If Not IsEmpty(item) Then
        text = text & item & Chr(13)
    Else: End If
Next item
If text <> "" Then
    text = Left(text, Len(text) - 1)
Else: End If
ArrayToStr = text

End Function


Function GetBaseData(data As Variant)
Const col_deno_sql As Integer = 2
Dim wb As String
Dim wb_path As String
Dim data_sql As Variant
Dim row_sql As Long
Dim sql_str As String
Dim index_hierarchy As String
Dim deno As String
Dim norm_array As Variant
Dim exist As Boolean
Dim row As Long
Dim norm As String
Dim product_type As String

sql_str = "SELECT * FROM [Таблица$]"

wb_path = Paths.NormAllPath
wb = Paths.NormAllName


data_sql = SQL.SqlSelect(sql_str, wb, wb_path, "no")

For row = UBound(data) To 1 Step -1
    data = TypeOfProduct.GetBaseValues(data, row)
    deno = data(row, Calculation.col_deno)
        If deno <> "" And Not IsEmpty(data(row, Calculation.col_type)) Then
            norm = ""
            product_type = data(row, Calculation.col_type)
            exist = False
            For row_sql = LBound(data_sql, 2) To UBound(data_sql, 2) ' Step -1
                If deno = data_sql(col_deno_sql, row_sql) Then
                    data = GetOperationsDataByType(data, data_sql, row_sql, row, product_type)
'                    If IsNumeric(data(row, Calculation.col_new_one_calc)) And data(row, Calculation.col_new_one_calc) > 0 Then
'                        data(row, Calculation.col_new_all_calc) = data(row, Calculation.col_new_one_calc) * data(row, Calculation.col_num)
'                    End If
                    Exit For
                Else: End If
            Next row_sql
        Else: End If
Next row
GetBaseData = data

End Function


Private Function GetOperationsDataByType(data As Variant, data_sql As Variant, row_sql As Long, row_data As Long, product_type As String)
Const l_col As Integer = 21
Const col_name As Integer = 1
Const col_min As Integer = 2
Const col_max As Integer = 3
Const indent As Integer = 3

Dim min As Double
Dim max As Double

Dim name As String
Dim name_data As String
Dim cfft_operation As Double
Dim ws As Worksheet
Dim rng As Range
Dim l_row As Long
Dim time As Variant
Dim rules_data As Variant

'On Error GoTo ShutDownMacro
'product_type = "Изделие"
Set ws = ThisWorkbook.Worksheets(product_type)
l_row = DocumentAttribute.LastRow(ws, 1)
error_val = "ОШИБКА"

With ws
    Set rng = .Range(.Cells(indent, 1), .Cells(l_row, l_col))
    rules_data = rng.value
    
    For row = LBound(rules_data) To UBound(rules_data)
        name = rules_data(row, col_name)
        
        For col_data = Calculation.col_def_one_calc To Calculation.col_new_one_calc
            If IsNumeric(data(row_data, col_data)) And Not IsEmpty(data(row_data, col_data)) Then
                name_data = data(1, col_data)
                If name = name_data Then
                    data(row_data, col_data) = 0
                    min = rules_data(row, col_min)
                    max = rules_data(row, col_max)
                    
                    For col = 4 To l_col Step 2
                        cfft_operation = rules_data(row, col)
                        If Not IsEmpty(cfft_operation) Then
                            name_operation = rules_data(row, col + 1)
                            If Not IsEmpty(name_operation) Then
                                For col_sql = col_start_sql To UBound(data_sql)
                                    If data_sql(col_sql, 0) = name_operation And Not IsNull(data_sql(col_sql, row_sql)) Then
                                        time = error_val
                                        On Error Resume Next
                                            time = CDec(data_sql(col_sql, row_sql))
                                        If time = error_val Then
                                            data(row_data, col_data) = time
                                        Else:
                                            data(row_data, col_data) = Round(CDec(data(row_data, col_data)) + CDec(time * cfft_operation), 2)
                                        End If
                                        
                                    End If
                                Next col_sql
                            End If
                        End If
                    Next col
                    
                    Select Case True
                    Case data(row_data, col_data) = error_val
                        data(row_data, col_data) = error_val
                    Case data(row_data, col_data) < min
                        data(row_data, col_data) = min
                    Case data(row_data, col_data) > max
                        data(row_data, col_data) = max
                    End Select
                    
                End If
            End If
        Next col_data
    Next row
End With
GetOperationsDataByType = data
Exit Function

ShutDownMacro:
Call Calculation.ShowAllSheets(False)
Call Screen.Events(True)
MsgBox ("Тип (" & product_type & ") отсутствует в списке типов изделий")
End

End Function
