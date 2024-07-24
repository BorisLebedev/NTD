Attribute VB_Name = "Time"
Sub Time()
Const index_col As Integer = 7
Dim data As Variant
Dim ws As Worksheet
Dim l_row As Long
Dim rng As Range
Dim rng_base As Range

Call Screen.Events(False)
Set ws = ActiveSheet()
With ws
    l_row = DocumentAttribute.LastRow(ws, Main.COL_NAME)
    If l_row > Main.TOP_INDENT Then
        Set rng = .Range(.Cells(Main.TOP_INDENT + 1, 1), .Cells(l_row, Main.L_COL))
    Else
        Set rng = .Range(.Cells(Main.TOP_INDENT + 1, 1), .Cells(Main.TOP_INDENT + 1, Main.L_COL))
    End If
    data = rng.value
    
    If MsgBox("Загрузить данные из таблицы трудоемкостей?", vbYesNo, "Confirm") = vbYes Then
        data = GetBaseData(data)
    Else
        Exit Sub
    End If
    
    rng = data
End With
Call Screen.Events(True)
End Sub



Private Function GetDenoFromCulc(data As Variant, row As Long)
Dim text As String
Dim deno As String

text = data(row, Calculation.COL_DENO)
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
'Const wb As String = "_Таблица трудоемкостей.xlsm"
'Const wb_path As String = "P:\20016000 Технологический отдел\Исходящие сл. зап. для ПЭО\"
'Const wb_path As String = "E:\Test Folder\Projects\НТД\"
Const col_name_sql As Integer = 1
Const col_deno_sql As Integer = 2
Const col_time_sql As Integer = 3
'Dim wb As String
Dim wb_path As String
Dim wb As String
Dim data_sql As Variant
Dim row_sql As Long
Dim sql_str As String
Dim index_hierarchy As String
Dim deno As String
Dim name As String
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
    deno = data(row, Main.COL_DENO)
        If deno <> "" Then
            norm = ""
            For row_sql = LBound(data_sql, 2) To UBound(data_sql, 2) ' Step -1
                If deno = data_sql(col_deno_sql, row_sql) Then
                    data(row, Main.COL_TIME) = Round(CDec(data_sql(col_time_sql, row_sql)), 2)
                    
'                    data = GetOperationsDataByType(data, data_sql, row_sql, row, product_type)
'                    If IsNumeric(data(row, Calculation.col_new_one_calc)) And data(row, Calculation.col_new_one_calc) > 0 Then
'                        data(row, Calculation.col_new_all_calc) = data(row, Calculation.col_new_one_calc) * data(row, Calculation.col_num)
'                    End If
                    Exit For
                Else: End If
            Next row_sql
        Else
            name = data(row, Main.COL_NAME)
            norm = ""
            For row_sql = LBound(data_sql, 2) To UBound(data_sql, 2) ' Step -1
                If deno = data_sql(col_name_sql, row_sql) Then
                    data(row, Main.COL_TIME) = Round(CDec(data_sql(col_time_sql, row_sql)), 2)
                    
'                    data = GetOperationsDataByType(data, data_sql, row_sql, row, product_type)
'                    If IsNumeric(data(row, Calculation.col_new_one_calc)) And data(row, Calculation.col_new_one_calc) > 0 Then
'                        data(row, Calculation.col_new_all_calc) = data(row, Calculation.col_new_one_calc) * data(row, Calculation.col_num)
'                    End If
                    Exit For
                Else: End If
            Next row_sql
        
        End If
Next row
GetBaseData = data

End Function



Function GetSubArray(data As Variant, column As Integer)
Dim data_calc As Variant

ReDim data_calc(LBound(data) To UBound(data), 1 To 1)

For row = LBound(data) To UBound(data)
    data_calc(row, 1) = data(row, column)
Next row
GetSubArray = data_calc

End Function
