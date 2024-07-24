Attribute VB_Name = "Based"
Sub Main()
Const index_col As Integer = 7
Dim data As Variant
Dim ws As Worksheet
Dim l_row As Long
Dim rng As Range
Dim rng_base As Range

Call Screen.Events(False)
Set ws = ActiveSheet()
With ws
    l_row = DocumentAttribute.LastRow(ws, Calculation.col_name) + 1
    If l_row > Calculation.top_indent Then
        Set rng = .Range(.Cells(Calculation.top_indent + 1, 1), .Cells(l_row, Calculation.l_col))
    Else
        Set rng = .Range(.Cells(Calculation.top_indent + 1, 1), .Cells(Calculation.top_indent + 1, Calculation.l_col))
    End If
    data = rng.value
    
    data = GetNewBaseData(data)
    
'    If MsgBox("Загрузить данные из новой таблицы трудоемкостей?", vbYesNo, "Confirm") = vbYes Then
'        data = GetNewBaseData(data)
'    Else
'        data = GetBaseData(data)
'    End If
    
    data_base = GetSubArray(data, Calculation.col_base)
    Set rng_base = .Range(.Cells(Calculation.top_indent + 1, Calculation.col_base), .Cells(l_row, Calculation.col_base))
    rng_base = data_base
    Call Calculation.RowHeights(rng, data)
End With
Call Screen.Events(True)
End Sub


Private Function GetBaseData(data As Variant)
Const wb As String = "БАЗА НОРМИРОВАНИЯ.xlsm"
Const wb_path As String = ""
Const col_deno_sql As Integer = 1
Const col_norm_sql As Integer = 2
Const col_prim_sql As Integer = 3
Const col_prod_sql As Integer = 4
Const col_date_sql As Integer = 5
Const col_empl_sql As Integer = 6
Const col_file_sql As Integer = 7
Dim data_sql As Variant
Dim sql_str As String
Dim index_hierarchy As String
Dim deno As String
Dim norm_array As Variant
Dim exist As Boolean
Dim row As Long

Dim norm As String

sql_str = "SELECT F1, F2, F3, F4, F5, F6, F7, F8 FROM [Лист1$]"

data_sql = SQL.SqlSelect(sql_str, wb, wb_path, "no")

For row = UBound(data) To 1 Step -1
    index_hierarchy = data(row, Calculation.col_hierarchy)
    If index_hierarchy <> "" Then
        deno = GetDenoFromCulc(data, row)
        If deno <> "" Then
            ReDim norm_array(1 To 1)
            norm = ""
            exist = False
            For row_sql = UBound(data_sql, 2) To LBound(data_sql, 2) Step -1
                If InStr(1, data_sql(col_deno_sql, row_sql), deno, vbTextCompare) Then
                    norm = data_sql(col_norm_sql, row_sql) & " | " & data_sql(col_date_sql, row_sql) & " | " & _
                           data_sql(col_deno_sql, row_sql) & " | " & data_sql(col_file_sql, row_sql)
                    exist = ExistInArray(norm_array, norm)
                    If Not exist Then
                        norm_array = AddToArray(norm_array, norm)
                    Else: End If
                Else: End If
            Next row_sql
            data(row, Calculation.col_base) = ArrayToStr(norm_array)
        Else: End If
    Else: End If
Next row
GetBaseData = data

End Function


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


Private Function GetNewBaseData(data As Variant)
'Const wb As String = "_Таблица трудоемкостей.xlsm"
'Const wb_path As String = ""
Const col_deno_sql As Integer = 2
Const col_norm_sql As Integer = 3
Const col_date_sql As Integer = 4
Const col_empl_sql As Integer = 5
Const col_prod_sql As Integer = 6
Const col_prim_sql As Integer = 8
Dim wb As String
Dim wb_path As String
Dim data_sql As Variant
Dim sql_str As String
Dim index_hierarchy As String
Dim deno As String
Dim norm_array As Variant
Dim exist As Boolean
Dim row As Long
Dim norm As String

wb_path = Paths.NormAllPath
wb = Paths.NormAllName
sql_str = "SELECT F1, F2, F3, F4, F5, F6, F7, F8, F9 FROM [Таблица$]"

data_sql = SQL.SqlSelect(sql_str, wb, wb_path, "no")

'For row = UBound(data) To 1 Step -1

For row = 1 To UBound(data)
    index_hierarchy = data(row, Calculation.col_hierarchy)
    If index_hierarchy <> "" Then
        deno = GetDenoFromCulc(data, row)
        If deno <> "" Then
            ReDim norm_array(1 To 1)
            norm = ""
            exist = False
            For row_sql = LBound(data_sql, 2) To UBound(data_sql, 2)
                If InStr(1, data_sql(col_deno_sql, row_sql), deno, vbTextCompare) <> 0 Then
                    norm = data_sql(col_norm_sql, row_sql) & "  |  " & data_sql(col_date_sql, row_sql) & "  |  " & _
                           data_sql(col_deno_sql, row_sql) & "  |  " & data_sql(col_prod_sql, row_sql) & "  |  " & data_sql(col_empl_sql, row_sql)
                    exist = ExistInArray(norm_array, norm)
                    If Not exist Then
                        norm_array = AddToArray(norm_array, norm)
                    Else: End If
                Else: End If
            Next row_sql
            data(row, Calculation.col_base) = ArrayToStr(norm_array)
        Else: End If
    Else: End If
Next row
GetNewBaseData = data

End Function
