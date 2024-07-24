Attribute VB_Name = "DataMK"
Const COL_DNKD As Integer = 1
Const COL_DNTD As Integer = COL_DNKD + 1
Const COL_ONUM As Integer = COL_DNTD + 1
Const COL_OPER As Integer = COL_ONUM + 1
Const COL_NORM As Integer = COL_OPER + 1
Const COL_PATH As Integer = COL_NORM + 1
Global Const TOP_INDENT As Integer = 2
'Const DOC_PATH As String = ""
Const F_COL As Integer = COL_DNKD
Global Const L_COL As Integer = COL_PATH


Sub main()
Dim data As Variant
Dim rng As Range
Dim ws As Worksheet

Set ws = ThisWorkbook.Worksheets("Данные из МК")
Call TimeCollector.ClearWS(ws)

With ws
    data = GetAllData()
    Set rng = .Range(.Cells(TOP_INDENT + 1, 1), .Cells(TOP_INDENT + 1 + UBound(data), L_COL))
    rng = data
End With

With ws
    Set rng = .Range(.Cells(TOP_INDENT, 1), .Cells(TOP_INDENT + 1 + UBound(data), L_COL))
    rng.AutoFilter
End With

With ws.Sort
    .SortFields.Add Key:=ws.Cells(TOP_INDENT + 1, COL_DNTD), order:=xlAscending
    .SortFields.Add Key:=ws.Cells(TOP_INDENT + 1, COL_ONUM), order:=xlAscending
    .SetRange rng
    .header = xlYes
    .Apply
End With

With ws
    Set rng = .Range(.Cells(TOP_INDENT + 1, 1), .Cells(TOP_INDENT + 1 + UBound(data), L_COL))
    Call SameProductsSameColor(rng)
End With

rng.Borders.LineStyle = XlLineStyle.xlContinuous
rng.RowHeight = 15
Call SetHeaderData(ws)
End Sub


Sub SameProductsSameColor(rng As Range)
Const color1 As Integer = 19
Const color2 As Integer = 2
Const color3 As Integer = 35    'green
Const color4 As Integer = 3     'red
Const color5 As Integer = 34    'blue
Const color6 As Integer = 40    'brown
Dim dublicate As Boolean
Dim color As Integer
Dim deno As String
Dim deno_next As String
Dim row As Long
Dim norm
Dim norm_next


color = color1
deno = rng.Rows(1).Cells(1, COL_DNTD).value
rng.Rows(1).Interior.ColorIndex = color

For row = 2 To rng.Rows.Count
    deno_next = rng.Rows(row).Cells(1, COL_DNTD).value
    index_next = rng.Rows(row).Cells(1, COL_ONUM).value
    
    If deno <> "" Then
        If deno <> deno_next Then
            dublicate = False
            If color = color1 Then
                color = color2
            Else
                color = color1
            End If
            norm_sum = 0
            rng.Rows(row).Interior.ColorIndex = color
        Else
            rng.Rows(row).Interior.ColorIndex = color
            norm = GetNum(rng.Rows(row - 1).Cells(1, COL_NORM).Value2)
            index = rng.Rows(row - 1).Cells(1, COL_ONUM).Value2
            index_next = rng.Rows(row).Cells(1, COL_ONUM).Value2
            
            If index = index_next Then
                dublicate = True
                If index = "В" Then
                    color_dub = color6
                Else
                    color_dub = color5
                End If
            Else: End If
            
            If dublicate Then
                rng.Rows(row - 1).Cells(1, COL_PATH).Interior.ColorIndex = color_dub
                rng.Rows(row).Cells(1, COL_PATH).Interior.ColorIndex = color_dub
            Else: End If
            
            norm_sum = norm_sum + norm
            If index_next = "С" Then
                norm_next = GetNum(rng.Rows(row).Cells(1, COL_NORM).Value2)
                If norm_sum = norm_next Then
                    rng.Rows(row).Cells(1, COL_NORM).Interior.ColorIndex = color3
                    rng.Rows(row).Cells(1, COL_OPER).Interior.ColorIndex = color3
                    rng.Rows(row).Cells(1, COL_ONUM).Interior.ColorIndex = color3
                Else
                    rng.Rows(row).Cells(1, COL_NORM).Interior.ColorIndex = color4
                    rng.Rows(row).Cells(1, COL_OPER).Interior.ColorIndex = color4
                    rng.Rows(row).Cells(1, COL_ONUM).Interior.ColorIndex = color4
                End If
            Else: End If
            
        End If
    Else: End If
    deno = deno_next
Next row
End Sub


Private Function GetNum(val As Variant)
Dim result

On Error GoTo EXIT_FUNC
GetNum = CDec(val)
Exit Function

EXIT_FUNC:
GetNum = 0

End Function


Private Sub SetHeaderData(ws As Worksheet)
Dim rng As Range

With ws
    
    .Cells(1, COL_DNKD) = "Обозначение КД"
    .Cells(1, COL_DNTD) = "Обозначение ТД"
    .Cells(1, COL_ONUM) = "№"
    .Cells(1, COL_OPER) = "Наименование"
    .Cells(1, COL_NORM) = "Тр-ть"
    .Cells(1, COL_NORM) = "Наименование файла"
    
    Set rng = .Range(.Cells(1, 1), .Cells(1, L_COL))
    rng.RowHeight = 30
    
    Set rng = .Range(.Cells(2, 1), .Cells(2, L_COL))
    rng.RowHeight = 13
    
    Set rng = .Range(.Cells(1, 1), .Cells(TOP_INDENT, L_COL))
    rng.Borders.LineStyle = XlLineStyle.xlContinuous
    
End With

End Sub


Private Function GetAllData()
Dim data As Variant
Dim counter As Integer
Dim doc_counter As Integer
Dim wb_name As String
Dim is_mk As Boolean
Dim doc_type As String

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFolder = oFSO.GetFolder(DOC_PATH)
 
counter = 1
doc_counter = FileCounter(DOC_PATH)
For Each oFile In oFolder.Files
    wb_name = oFile.name
    doc_type = DocTypeByName(wb_name)
    If doc_type = "МК" Then
        If InStr(wb_name, "$") = 0 Then
            Application.StatusBar = format(((counter / doc_counter) * 100), "#,#0") & "%..." & wb_name
            data = GetData(data, wb_name, DOC_PATH)
            counter = counter + 1
        Else: End If
    Else
        wb_path_new = "E:\Test Folder\Projects\STC_DB_N\_Общая ТД (Проблемы с форматом)\"
        FileCopy DOC_PATH & wb_name, wb_path_new & wb_name
        Kill DOC_PATH & wb_name
    End If
Next oFile
Application.StatusBar = ""
data = DocumentAttribute.Transpose2dArray(data)
GetAllData = data
End Function


Private Function DocTypeByName(wb_name As String)
Const DENO_LENGHT As Integer = 16
Const SYMB As String = "("
Dim doc_type As String
Dim begin_symb As Integer
Dim deno As String

begin_symb = InStr(1, wb_name, SYMB, vbTextCompare)
If begin_symb <> 0 Then
    deno = Left(Right(wb_name, Len(wb_name) - begin_symb), DENO_LENGHT)
    doc_type = DocumentAttribute.TdTypeDoc(deno)
Else
    doc_type = ""
End If
DocTypeByName = doc_type
End Function


Private Function GetData(data As Variant, wb As String, wb_path As String)
Const R_DNKD As Integer = 9
Const C_TEXT_01 As Integer = 1
Const C_DNKD_40 As Integer = 16
Const C_DNKD_42 As Integer = 18

Const R_DNTD As Integer = R_DNKD
Const C_TEXT_02 As Integer = 2
Const C_DNTD_40 As Integer = 34
Const C_DNTD_42 As Integer = 36

Const R_NOMR_F As Integer = 18
Const C_NOMR_F_40 As Integer = 40
Const C_NOMR_F_42 As Integer = 40

Const C_INDX_44 As Integer = 0
Const C_ONUM_44 As Integer = 8
Const C_OPER_44 As Integer = 9
Const C_NORM_44 As Integer = 44
Const C_INDX_43 As Integer = 0
Const C_ONUM_43 As Integer = 7
Const C_OPER_43 As Integer = 8
Const C_NORM_43 As Integer = 43

Dim col_norm_mk As Integer
Dim sql_data As Variant
Dim sql_str As String
Dim start As Long
Dim operation_name As String
Dim dict As New Scripting.Dictionary
Dim dnkd As String
Dim dntd As String
Dim row As Integer
Dim norm_array As Variant
Dim row_without_merge As Integer

Dim f_sheet As Variant
Dim l_sheet As Variant

Dim f_sheet_columns As Variant
Dim l_sheet_columns As Variant

norm_sum = ""
Set dict = GetDictOfSheets(wb, wb_path)

f_sheet = GetFirstSheet(dict, wb, wb_path)
If Not IsEmpty(f_sheet) Then
    f_sheet_columns = GetLastCol(f_sheet, "2")
    Select Case f_sheet_columns
        Case 40
            dnkd = f_sheet(C_DNKD_40, R_DNKD)
            dntd = f_sheet(C_DNTD_40, R_DNTD)
            norm_sum = GetNorm(f_sheet, C_NOMR_F_40, R_NOMR_F, C_TEXT_01)
        Case 42
            dnkd = f_sheet(C_DNKD_42, R_DNKD)
            dntd = f_sheet(C_DNTD_42, R_DNTD)
            norm_sum = GetNorm(f_sheet, C_NOMR_F_42, R_NOMR_F, C_TEXT_02)
        Case Else
    End Select

    
'    sql_str = ""
    For Each Key In dict.Keys()
'        For Key = 2 To dict.Count()
        sql_str = ""
        If InStr(1, Key, "Форма 2", vbTextCompare) = 0 And Key <> "1" Then
'                sql_str = sql_str + "SELECT * FROM [" & Key & "$] WHERE F1 IS NOT NULL UNION ALL "
'                sql_str = sql_str + "SELECT * FROM [" & Key & "$] OUTER UNION CORRESPONDING "
'            Else: End If
            sql_str = sql_str + "SELECT * FROM [" & Key & "$]"
'        Next Key
'
'        sql_str = Left(sql_str, Len(sql_str) - 10)
'        sql_str = Left(sql_str, Len(sql_str) - Len(" OUTER UNION CORRESPONDING "))
            l_sheet = SQL.SqlSelect(sql_str, wb, wb_path)
        Else
            l_sheet = Empty
        End If
        
        If Not IsEmpty(l_sheet) Then
            l_sheet_columns = GetLastCol(l_sheet, "1б")
            Select Case l_sheet_columns
                Case 44
                    COL_ONUM_MK = C_ONUM_44
                    COL_OPER_MK = C_OPER_44
                    col_norm_mk = C_NORM_44
                Case 43
                    COL_ONUM_MK = C_ONUM_43
                    COL_OPER_MK = C_OPER_43
                    col_norm_mk = C_NORM_43
                Case Else
            End Select
            
            index = ""
            For row = 15 To UBound(l_sheet, 2)
                If norm_sum = "" And index = "" And row > 14 And row < 50 And Key = "2" Then
        '            norm_array = Array(l_sheet(COL_NORM_MK, row), _
        '                               l_sheet(COL_NORM_MK - 1, row), _
        '                               l_sheet(COL_NORM_MK - 2, row), _
        '                               l_sheet(COL_NORM_MK - 3, row), _
        '                               l_sheet(COL_NORM_MK - 4, row))
                    norm_sum = GetNorm(l_sheet, col_norm_mk, row)
                Else: End If
            
                index = GetIndex(l_sheet, row)
                row_without_merge = GetRowWithoutMerge(row)
    '            row_without_merge = row
                Select Case index
                
                Case "А"
                    data = AddRowToData(data)
                    data_row = UBound(data, 2)
                    data(COL_DNKD, data_row) = dnkd
                    data(COL_DNTD, data_row) = dntd
                    data(COL_ONUM, data_row) = GetOperationNum(l_sheet(COL_ONUM_MK, row_without_merge))
                    data(COL_OPER, data_row) = GetOperationName(l_sheet(COL_OPER_MK, row_without_merge), l_sheet(COL_ONUM_MK, row_without_merge))
                    data(COL_PATH, data_row) = wb
                Case "Б"
        '            norm_array = Array(l_sheet(COL_NORM_MK, row), _
        '                               l_sheet(COL_NORM_MK - 1, row), _
        '                               l_sheet(COL_NORM_MK - 2, row), _
        '                               l_sheet(COL_NORM_MK - 3, row), _
        '                               l_sheet(COL_NORM_MK - 4, row))
                    data(COL_NORM, data_row) = GetNorm(l_sheet, col_norm_mk, row_without_merge)
                End Select
            Next row
        Else: End If
    Next Key
    
    data = AddRowToData(data)
    data_row = UBound(data, 2)
    data(COL_DNKD, data_row) = dnkd
    data(COL_DNTD, data_row) = dntd
    data(COL_ONUM, data_row) = "С"
    data(COL_OPER, data_row) = "Сумма"
    data(COL_NORM, data_row) = norm_sum
    data(COL_PATH, data_row) = wb
    
Else: End If
GetData = data
End Function


Private Function GetRowWithoutMerge(row)

If row Mod 2 = 0 Then
    GetRowWithoutMerge = row
Else
    GetRowWithoutMerge = row + 1
End If

End Function


Private Function GetNormColumn(dntd As String)

sql_columns = "F1=0, F2=0, F9, F10, F44, F45"

'If Left(Right(dntd, 5), 1) = "9" Then
'    sql_columns = "F1=0, F2=0, F9, F10, F44, F45=0"
'Else
'    sql_columns = "F1=0, F2=0, F9, F10, F44, F45"
'End If

GetNormColumn = sql_columns
End Function


Private Function GetDictOfSheets(wb As String, wb_path As String) As Dictionary
Dim CON As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim sheets As Variant
Dim name As String
Dim dict As New Scripting.Dictionary
Dim arr As Variant
Dim num As Integer

CON.Provider = "Microsoft.ACE.OLEDB.12.0"
CON.ConnectionString = "data source=" & wb_path & wb & "; extended properties=""Excel 12.0 xml;HDR=no"""
CON.Open
Set sheets = CON.OpenSchema(adSchemaTables)
Set dict = CreateObject("Scripting.Dictionary")
arr = Array()
Do While Not sheets.EOF
    name = sheets.Fields("table_name").value
    If Left(name, 4) <> "Лист" And InStr(name, "#") = 0 Then
        If InStr(name, "$") <> 0 Then
            name = Left(name, InStr(name, "$") - 1)
            Do While InStr(name, "'") <> 0
                name = Replace(name, "'", "")
            Loop
            dict(name) = 1
            
        Else: End If
    Else: End If
    sheets.MoveNext
Loop
CON.Close



Set GetDictOfSheets = dict
End Function


Private Function GetMainDataArray(dict As Dictionary, wb As String, wb_path As String)
Dim sql_str As String

sql_str = sql_str + "SELECT F17, F35 FROM [Форма 2$] WHERE F1 = 'Разработал'"
sql_data = SQL.SqlSelect(sql_str, wb, wb_path)

GetMainDataArray = sql_data
End Function


Private Function GetSumNorm(dict As Dictionary, wb As String, wb_path As String)
Dim sql_str As String

sql_str = "SELECT F39, F40, F41 FROM [Форма 2$] WHERE F1 = '01'"
sql_data = SQL.SqlSelect(sql_str, wb, wb_path)

GetSumNorm = sql_data
End Function


Private Function GetFirstSheet(dict As Dictionary, wb As String, wb_path As String)
Dim sql_str As String

For Each Key In dict.Keys()
    If InStr(1, Key, "Форма 2", vbTextCompare) <> 0 Or Key = "1" Then
        'sql_str = sql_str + "SELECT " & sql_columns & " FROM [" & Key & "$] WHERE F9 IS NOT null AND F2 IS NOT null UNION ALL "
        sql_str = "SELECT * FROM [" & Key & "$]"
        sql_data = SQL.SqlSelect(sql_str, wb, wb_path)
    Else: End If
Next Key


GetFirstSheet = sql_data
End Function


Private Function GetNorm(data As Variant, col_norm_mk As Integer, row As Integer, Optional first_page As Integer = 0)
Dim norm_array As Variant

norm_array = Array(data(col_norm_mk, row), _
                   data(col_norm_mk - 1, row), _
                   data(col_norm_mk - 2, row), _
                   data(col_norm_mk - 3, row), _
                   data(col_norm_mk - 4, row), _
                   data(col_norm_mk, row - 2), _
                   data(col_norm_mk - 1, row - 2), _
                   data(col_norm_mk - 2, row - 2), _
                   data(col_norm_mk - 3, row - 2), _
                   data(col_norm_mk - 4, row - 2))

For Each Item In norm_array
    If Not IsNull(Item) And Not IsEmpty(Item) And Item <> "" Then
        On Error Resume Next
        result = CDec(Item)
    Else: End If
Next Item

If IsEmpty(result) And first_page <> 0 Then
    On Error Resume Next
    result = CDec(data(first_page, row))
Else: End If

GetNorm = result
End Function


Private Function GetLastCol(sheet_array As Variant, form As String)
Const row As Integer = 6
Dim L_COL As Integer

For col = UBound(sheet_array) To 0 Step -1
    If sheet_array(col, row) = "Дата" Then
        Select Case form
        Case "2"
            L_COL = col + 1
        Case "1б"
            L_COL = col
        End Select
        Exit For
    Else: End If
Next col

GetLastCol = L_COL
End Function


Private Function AddRowToData(data As Variant)

If IsEmpty(data) Then
    ReDim data(F_COL To L_COL, 0 To 0)
Else
    ReDim Preserve data(LBound(data) To UBound(data), _
                        LBound(data, 2) To UBound(data, 2) + 1)
End If
AddRowToData = data
End Function


Private Function GetIndex(sheet As Variant, row As Integer)
Dim text_p1 As String
Dim text_p2 As String
Dim text As String
Dim index As String

text = ""
text_next = ""
text = ToStr(sheet(0, row)) & ToStr(sheet(1, row))
On Error Resume Next
text_next = ToStr(sheet(0, row + 1)) & ToStr(sheet(1, row))

Select Case text
Case "Цех", "Код, наименование оборудования", "Наименование детали, сб. единицы или материала"
    index = ""
Case ""
    index = ""
Case Else
    If text_next <> text Then
        index = ToStr(sheet(0, row))
    Else: End If
End Select

If Not IsNull(index) Then
    index = Left(index, 1)
Else: End If

GetIndex = index
End Function


Private Function ToStr(val As Variant)
Dim new_val As String

Select Case True
Case IsNull(val)
    new_val = ""
Case IsEmpty(val)
    new_val = ""
Case Else
    new_val = CStr(val)
End Select

ToStr = new_val
End Function



Private Function GetOperationNum(val As Variant)
Dim num As String

If IsNull(val) Then
    num = "В"
Else
    num = CInt(val) / 5
End If

GetOperationNum = num
End Function


Private Function GetOperationName(name As Variant, num As Variant)
Dim new_name As String

If IsNull(num) And IsNull(name) Then
    new_name = "Сборка и монтаж ИЭТ"
Else
    new_name = name
End If

GetOperationName = new_name
End Function
