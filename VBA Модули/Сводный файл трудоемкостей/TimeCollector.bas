Attribute VB_Name = "TimeCollector"
'Global OPERATIONS_CORRECTION As Variant
'Global Const CALCULATION_PATH As String = ""
Global Const COL_RNUM As Integer = 1
Global Const COL_NAME As Integer = COL_RNUM + 1
Global Const COL_DENO As Integer = COL_NAME + 1
Global Const COL_NORM As Integer = COL_DENO + 1
Global Const COL_DATE As Integer = COL_NORM + 1
Global Const COL_EFIO As Integer = COL_DATE + 1
Global Const COL_PROJ As Integer = COL_EFIO + 1
Global Const COL_LINK As Integer = COL_PROJ + 1
Global Const COL_OPER As Integer = COL_LINK + 1

Global Const TOP_INDENT As Integer = 3
Const index_col As Integer = COL_NAME
Dim L_COL As Integer

Global Const COL_LEVEL_CALC As Integer = 0
Global Const COL_HIERARCHY_CALC As Integer = COL_LEVEL_CALC + 1
Global Const COL_NAME_CALC As Integer = COL_HIERARCHY_CALC + 1
Global Const COL_DENO_CALC As Integer = COL_NAME_CALC + 1
Global Const COL_DENO_TD_CALC As Integer = COL_DENO_CALC + 1
Global Const COL_OPER_CALC As Integer = COL_NAME_CALC
Global Const COL_NUM_CALC As Integer = COL_DENO_TD_CALC + 1
Global Const COL_NORM_CALC As Integer = COL_NUM_CALC + 1
Global Const COL_DATE_CALC As Integer = COL_NORM_CALC + 1
Global Const COL_EFIO_CALC As Integer = COL_DATE_CALC + 1
Global Const COL_PROJ_CALC As Integer = COL_EFIO_CALC + 1
Global Const COL_CFIX_CALC As Integer = COL_PROJ_CALC + 1
Global Const COL_COMM_CALC As Integer = COL_CFIX_CALC + 1



Sub main()
Dim data As Variant
Dim ws As Worksheet
Dim ws_operation As Worksheet
Dim operations_num As Integer
Dim l_row As Long
Dim rng As Range
Dim data_operations As Variant
Dim data_corrections As Variant


Call Screen.Events(False)
Calculation.OPERATIONS = Calculation.GetOperationArray()
Calculation.OPERATIONS_CORRECTION = Calculation.GetOperationCorrectionArray()

Set ws = ThisWorkbook.Worksheets("Таблица")
Set ws_operation = ThisWorkbook.Worksheets("Операции")
'Set ws_correction = ThisWorkbook.Worksheets("Исправления")

Call ClearWS(ws)
operations_num = DocumentAttribute.LastRow(ws_operation, 1)
L_COL = COL_OPER + operations_num
Call SetHeaderData(ws, ws_operation, operations_num)

With ws
    data = GetAllData()
    data = DocumentAttribute.Transpose2dArray(data)
    data = GetProducts(data)
    Set rng = .Range(.Cells(TOP_INDENT + 1, 1), .Cells(TOP_INDENT + UBound(data), L_COL))
End With

rng = data
ReDim arr(0 To L_COL - 1)
For i = 0 To L_COL - 1
    arr(i) = i + 1
Next i

Call RemoveDuplicates(rng)

rng.Borders.LineStyle = XlLineStyle.xlContinuous
rng.RowHeight = 15

With ws.Sort
    .SortFields.Add Key:=ws.Cells(TOP_INDENT + 1, COL_DENO), order:=xlAscending
    .SortFields.Add Key:=ws.Cells(TOP_INDENT + 1, COL_DATE), order:=xlDescending
    .SortFields.Add Key:=ws.Cells(TOP_INDENT + 1, COL_NAME), order:=xlAscending
    .SetRange rng
    .header = xlYes
    .Apply
    Call SameProductsSameColor(rng)
End With

Call HideEmptyColumns(rng)
ThisWorkbook.Names.Add name:="ALL_DATA", RefersTo:="=Нормы!$A$3:$H$" & (UBound(data))

With ws
    Set rng = .Range(.Cells(TOP_INDENT, 1), .Cells(UBound(data), L_COL))
    rng.AutoFilter
End With

Call Screen.Events(True)
Application.StatusBar = "Обновление завершено"
End Sub


Private Sub RemoveDuplicates(rng As Range)
Dim intArray As Variant
Dim i As Integer

With rng
    ReDim intArray(0 To .Columns.Count - 1)
    For i = 0 To UBound(intArray)
        intArray(i) = i + 1
    Next i
    .RemoveDuplicates Columns:=(intArray), header:=xlNo
End With

End Sub


Sub ClearWS(ws As Worksheet)

With ws
    .AutoFilter.ShowAllData
    .Sort.SortFields.Clear
    .Cells.EntireColumn.Hidden = False
    .Cells.UnMerge
    .Cells.ClearContents
    .Cells.Interior.ColorIndex = 2
    .Cells.Borders.LineStyle = None
End With

End Sub


Private Function HideEmptyColumns(rng As Range)
Dim not_empty As Boolean

For col = 1 To rng.Columns.Count
    data = rng.Columns(col)
    not_empty = False
    For row = LBound(data) To UBound(data)
        If Not IsEmpty(data(row, 1)) Then
            not_empty = True
            Exit For
        Else: End If
    Next row
    
    If Not not_empty Then
        rng.Cells(1, col).EntireColumn.Hidden = True
    Else: End If
        
Next col
rng.Cells(1, COL_RNUM).EntireColumn.Hidden = True
End Function


Private Sub SetHeaderData(ws As Worksheet, ws_operation As Worksheet, operations_num As Integer)
Dim rng As Range

With ws
    
    .Cells(1, COL_NAME) = "Наименование"
    .Cells(1, COL_DENO) = "Обозначение КД"
    .Cells(1, COL_NORM) = "Тр-ть"
    .Cells(1, COL_DATE) = "Дата"
    .Cells(1, COL_EFIO) = "ФИО"
    .Cells(1, COL_PROJ) = "Проект"
    .Cells(1, COL_LINK) = "Ссылка"
    .Cells(1, COL_OPER) = "Операции"
    
    .Range(.Cells(1, COL_NAME), .Cells(2, COL_NAME)).Merge
    .Range(.Cells(1, COL_DENO), .Cells(2, COL_DENO)).Merge
    .Range(.Cells(1, COL_NORM), .Cells(2, COL_NORM)).Merge
    .Range(.Cells(1, COL_DATE), .Cells(2, COL_DATE)).Merge
    .Range(.Cells(1, COL_EFIO), .Cells(2, COL_EFIO)).Merge
    .Range(.Cells(1, COL_PROJ), .Cells(2, COL_PROJ)).Merge
    .Range(.Cells(1, COL_LINK), .Cells(2, COL_LINK)).Merge
    .Range(.Cells(1, COL_OPER), .Cells(2, COL_OPER)).Merge
    
    
    For row = 1 To operations_num
        .Cells(1, COL_OPER + row) = ws_operation.Cells(row, 1)
        .Cells(1, COL_OPER + row).Orientation = xlUpward
        .Cells(1, COL_OPER + row).VerticalAlignment = xlBottom
        .Cells(1, COL_OPER + row).ColumnWidth = 7
        .Range(.Cells(1, COL_OPER + row), .Cells(2, COL_OPER + row)).Merge
    Next row
    
    Set rng = .Range(.Cells(1, 1), .Cells(1, L_COL))
    rng.RowHeight = 30
    
    Set rng = .Range(.Cells(2, 1), .Cells(2, L_COL))
    rng.RowHeight = 100
    
    Set rng = .Range(.Cells(3, 1), .Cells(3, L_COL))
    rng.RowHeight = 13
    
    Set rng = .Range(.Cells(1, 1), .Cells(TOP_INDENT, L_COL))
    rng.Borders.LineStyle = XlLineStyle.xlContinuous
    
End With

End Sub


Private Sub SameProductsSameColor(rng As Range)
Const color1 As Integer = 19
Const color2 As Integer = 2
Dim color As Integer
Dim deno As String
Dim deno_next As String
Dim row As Long
Dim norm
Dim norm_next



color = color1
deno = rng.Rows(1).Cells(1, COL_DENO).value
rng.Rows(1).Interior.ColorIndex = color

For row = 2 To rng.Rows.Count
    deno_next = rng.Rows(row).Cells(1, COL_DENO).value
    If deno <> "" Then
        If deno <> deno_next Then
            
            If color = color1 Then
                color = color2
            Else
                color = color1
            End If
            rng.Rows(row).Interior.ColorIndex = color
        Else
            rng.Rows(row).Interior.ColorIndex = color
            norm = CDec(rng.Rows(row - 1).Cells(1, COL_NORM).Value2)
            norm_next = CDec(rng.Rows(row).Cells(1, COL_NORM).Value2)
            
            If norm_next <> norm Then
                Call SameNormColor(rng, row)
            Else
                Call OperationsColor(rng, row)
            End If
        
        End If
    Else: End If
    deno = deno_next
Next row
End Sub


Sub OperationsColor(rng As Range, row As Long)
Const color As Integer = 3
Dim norm
Dim norm_next

For col = COL_OPER + 1 To rng.Columns.Count
    norm = CDec(rng.Rows(row - 1).Cells(1, col).Value2)
    On Error Resume Next
        norm_next = CDec(rng.Rows(row).Cells(1, col).Value2)
    If norm <> norm_next Then
        rng.Cells(row, COL_OPER) = "отличается"
        rng.Cells(row - 1, COL_OPER) = "отличается"
        Exit For
    Else: End If
Next col
End Sub

Sub SameNormColor(rng As Range, row As Long)
Dim proj As String
Dim proj_next As String
Dim norm_date As Date
Dim norm_date_next As Date
Const color As Integer = 3

proj_next = rng.Rows(row).Cells(1, COL_PROJ).Value2
proj = rng.Rows(row - 1).Cells(1, COL_PROJ).Value2
norm_date_next = rng.Rows(row).Cells(1, COL_DATE).Value2
norm_date = rng.Rows(row - 1).Cells(1, COL_DATE).Value2

If proj_next = proj And norm_date_next And norm_date_next = norm_date Then
    rng.Rows(row).Interior.ColorIndex = color
    rng.Rows(row - 1).Interior.ColorIndex = color
Else
    rng.Rows(row).Cells(1, COL_NORM).Interior.ColorIndex = color
    rng.Rows(row - 1).Cells(1, COL_NORM).Interior.ColorIndex = color
End If

End Sub



Private Function GetProducts(data)
Dim row As Long
Dim row_new As Long
Dim row_sub As Integer
Dim index_hierarchy As String
Dim num As Integer
Dim norm As Double
Dim level As Integer
Dim data_new As Variant
Dim OPERATIONS As Variant
Dim operation As String


ReDim data_new(1 To L_COL, 1 To 1)
l_row = UBound(data)
For row = LBound(data) To UBound(data)
    
    
    If Not IsNull(data(row, COL_HIERARCHY_CALC)) Then
        index_hierarchy = data(row, COL_HIERARCHY_CALC)
        
        If index_hierarchy <> "" Then
            row_new = UBound(data_new, 2)
            ReDim Preserve data_new(LBound(data_new) To UBound(data_new), _
                                    LBound(data_new, 2) To UBound(data_new, 2) + 1)
            data_new(COL_RNUM, row_new) = data(row, COL_COMM_CALC)
            data_new(COL_NAME, row_new) = data(row, COL_NAME_CALC)
            data_new(COL_DENO, row_new) = data(row, COL_DENO_CALC)
            data_new(COL_NORM, row_new) = data(row, COL_NORM_CALC)
            data_new(COL_DATE, row_new) = data(row, COL_DATE_CALC)
            data_new(COL_EFIO, row_new) = data(row, COL_EFIO_CALC)
            data_new(COL_PROJ, row_new) = data(row, COL_PROJ_CALC)
            data_new(COL_LINK, row_new) = ">>>"
            data_new(COL_OPER, row_new) = ""
            
        Else
            operation = Trim(data(row, COL_NAME_CALC))
            If Not IsEmpty(data(row, COL_NORM_CALC)) Then
                For op_row = 2 To UBound(Calculation.OPERATIONS)
                    If operation = Calculation.OPERATIONS(op_row, 1) Then
                        data_new(COL_OPER + op_row, row_new) = data_new(COL_OPER + op_row, row_new) + data(row, COL_NORM_CALC)
                        Exit For
                    Else: End If
                Next op_row
            Else: End If
        End If
    Else: End If
Next
data_new = DocumentAttribute.Transpose2dArray(data_new)

GetProducts = data_new

End Function


Private Function GetAllData()

Dim counter As Integer
Dim doc_counter As Integer
Dim wb_name As String
Dim wb_path As String

Set oFSO = CreateObject("Scripting.FileSystemObject")
wb_path = CALCULATION_PATH
Set oFolder = oFSO.GetFolder(wb_path)
 
counter = 1
doc_counter = FileCounter(wb_path)
For Each oFile In oFolder.Files
    wb_name = oFile.name
    If InStr(wb_name, "$") = 0 Then
        Application.StatusBar = format(((counter / doc_counter) * 100), "#,#0") & "%..." & wb_name
        data = GetData(data, "\" & wb_name, wb_path)
        counter = counter + 1
    Else: End If
Next oFile
GetAllData = data
End Function


Private Function GetData(data As Variant, wb As String, wb_path As String)
Dim sql_data As Variant
Dim sql_str As String
Dim start As Long
Dim calc_date As Date
Dim calc_efio As String
Dim calc_proj As String
Dim operation_name As String
Dim time As String

sql_str = "SELECT F1, F2, F3, F4, F5, F6, F7, F8, F9, F10, F11 FROM [Расшифровка$]"

'WHERE NOT (" & _
          "F1 IS null AND F2 IS null AND F3 IS null AND " & _
          "F4 IS null AND F5 IS null AND F6 IS null AND " & _
          "F7 IS null AND F8 IS null)
sql_data = SQL.SqlSelect(sql_str, wb, wb_path, "no")

If IsEmpty(data) Then
    start = 0
    ReDim data(COL_LEVEL_CALC To COL_COMM_CALC, _
               LBound(sql_data, 2) To UBound(sql_data, 2) + 1)
Else
    start = UBound(data, 2)
    ReDim Preserve data(LBound(data) To UBound(data), _
                        LBound(data, 2) To UBound(data, 2) + UBound(sql_data, 2) + 1)
End If

lenght = 0
lenght = InStr(1, wb, "_") - 1
temp = Left(wb, lenght)
calc_proj = Right(temp, Len(temp) - 1)

lenght = InStr(lenght + 2, wb, "_") - 1
temp = Left(wb, lenght)
temp = Right(temp, Len(temp) - Len(calc_proj) - 2)
calc_date = format(temp, "DD.MM.YYYY")

lenght = InStrRev(wb, ".") - 1
temp = Left(wb, lenght)
lenght = InStrRev(wb, "_")
calc_efio = Right(temp, Len(temp) - lenght)


For row = LBound(sql_data, 2) + 1 To UBound(sql_data, 2)
    If Not IsNull(sql_data(COL_NORM_CALC, row)) And Not IsEmpty(sql_data(COL_NORM_CALC, row)) And sql_data(COL_NORM_CALC, row) <> "" Then
        If sql_data(COL_HIERARCHY_CALC, row) <> "" Then
            calc_deno = sql_data(COL_DENO_CALC, row)
            If Not IsNull(calc_deno) Then
                calc_deno = Replace(calc_deno, " ", "")
            End If
            On Error Resume Next
                sql_data(COL_NORM_CALC, row) = CDec(sql_data(COL_NORM_CALC, row))
            
            data(COL_LEVEL_CALC, row + start) = sql_data(COL_LEVEL_CALC, row)
            data(COL_HIERARCHY_CALC, row + start) = sql_data(COL_HIERARCHY_CALC, row)
            data(COL_NAME_CALC, row + start) = sql_data(COL_NAME_CALC, row)
            data(COL_DENO_CALC, row + start) = calc_deno
            data(COL_NORM_CALC, row + start) = sql_data(COL_NORM_CALC, row)
            data(COL_DATE_CALC, row + start) = calc_date
            data(COL_EFIO_CALC, row + start) = calc_efio
            data(COL_PROJ_CALC, row + start) = calc_proj
            data(COL_COMM_CALC, row + start) = row
            If data(COL_DENO_CALC, row + start) = "УИЕС.461434.001" Then
                data(COL_DENO_CALC, row + start) = "УИЕС.461434.001"
            Else: End If
        Else:
            
            
        
            If Not IsNull(sql_data(COL_NAME_CALC, row)) Then
                operation_name = sql_data(COL_NAME_CALC, row)
                operation_name = GetCorrectOperation(operation_name)
                
                If operation_name <> "" Then
                    data(COL_NAME_CALC, row + start) = GetCorrectOperation(operation_name)
                    On Error Resume Next
                        sql_data(COL_NORM_CALC, row) = CDec(sql_data(COL_NORM_CALC, row))
                    data(COL_NORM_CALC, row + start) = sql_data(COL_NORM_CALC, row)
'                    data(COL_NORM_CALC, row + start) = sql_data(COL_NORM_CALC, row)
                Else: End If
            Else: End If
        End If
    Else: End If
Next row

GetData = data

End Function


Function GetCorrectOperation(operation_name As String)
Dim operation As String

operation = ""

For row = LBound(Calculation.OPERATIONS) To UBound(Calculation.OPERATIONS)
    If operation_name = Calculation.OPERATIONS(row, 1) Then
        operation = Calculation.OPERATIONS(row, 1)
        Exit For
    Else: End If
Next row

If operation = "" Then
    For row = LBound(Calculation.OPERATIONS_CORRECTION) To UBound(Calculation.OPERATIONS_CORRECTION)
        If operation_name = Calculation.OPERATIONS_CORRECTION(row, 1) Then
            operation = Calculation.OPERATIONS_CORRECTION(row, 2)
            Exit For
        Else: End If
    Next row
Else: End If

GetCorrectOperation = operation
End Function


Sub CalculationLink(row As Long)
Dim date_in_wb_name As String
Dim proj_in_wb_name As String
Dim efio_in_wb_name As String
Dim wb_name As String
Dim row_num As Integer

proj_in_wb_name = Cells(row, COL_PROJ)
efio_in_wb_name = Cells(row, COL_EFIO)
date_in_wb_name = CStr(format(Cells(row, COL_DATE).Value2, "YYYY.MM.DD"))
row_num = Cells(row, COL_RNUM) + 1

wb_name = proj_in_wb_name & "_" & _
          date_in_wb_name & "_" & _
          efio_in_wb_name & ".xlsm"
          
ActiveWorkbook.FollowHyperlink (CALCULATION_PATH & wb_name)
ActiveWorkbook.Worksheets("Расшифровка").Activate
ActiveWorkbook.Worksheets("Расшифровка").Range("A" & row_num).Select

End Sub
