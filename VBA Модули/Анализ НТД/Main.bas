Attribute VB_Name = "main"
'Const
Global Const TOP_INDENT As Integer = 2
'Global Const FOLDER_ANALISYS As String = "\НТД для анализа\"


Global Const ROW_START As Integer = 12
Global Const COL_HIER_CALC As Integer = 0
Global Const COL_NAME_CALC As Integer = COL_HIER_CALC + 1
Global Const COL_DENO_CALC As Integer = COL_NAME_CALC + 1
Global Const COL_NUM_CALC As Integer = COL_DENO_CALC + 1
Global Const COL_MSR_CALC As Integer = COL_NUM_CALC + 1
Global Const COL_DEF_CALC As Integer = COL_MSR_CALC + 1
Global Const COL_DIS_CALC As Integer = COL_DEF_CALC + 1
Global Const COL_ASL_CALC As Integer = COL_DIS_CALC + 1
Global Const COL_REP_CALC As Integer = COL_ASL_CALC + 1
Global Const COL_RPR_CALC As Integer = COL_REP_CALC + 1
Global Const COL_TUN_CALC As Integer = COL_RPR_CALC + 1
Global Const COL_MAN_CALC As Integer = COL_TUN_CALC + 1
Global Const COL_TYPE_CALC As Integer = COL_MAN_CALC + 1


Global Const COL_HIER As Integer = 1
Global Const COL_NAME As Integer = COL_HIER + 1
Global Const COL_DENO As Integer = COL_NAME + 1
Global Const COL_NUM As Integer = COL_DENO + 1
Global Const COL_MSR As Integer = COL_NUM + 1

Global Const COL_DEF As Integer = COL_MSR + 1
Global Const COL_DIS As Integer = COL_DEF + 1
Global Const COL_ASL As Integer = COL_DIS + 1
Global Const COL_REP As Integer = COL_ASL + 1
Global Const COL_RPR As Integer = COL_REP + 1
Global Const COL_TUN As Integer = COL_RPR + 1
Global Const COL_MAN As Integer = COL_TUN + 1
Global Const COL_TIME As Integer = COL_MAN + 1
Global Const COL_TYPE As Integer = COL_TIME + 1

Global Const COL_PROD As Integer = COL_TYPE + 1
Global Const COL_LINK_H As Integer = COL_PROD + 1
Global Const COL_LINK As Integer = COL_LINK_H + 1

Global Const L_COL As Integer = COL_LINK



Sub Main()
Dim ws As Worksheet
Dim rng As Range
Dim last_row As Long

Call Screen.Events(False)
Set ws = ThisWorkbook.Worksheets("Таблица")

With ws
    row_data = GetAllData()
    If IsEmpty(row_data) Then
        GoTo EXIT_SUB
    Else: End If
    data = DocumentAttribute.Transpose2dArray(row_data)
    Call ClearWS(ws)
    Call SetHeaderData(ws)
    Set rng = .Range(.Cells(TOP_INDENT + 1, 1), .Cells(TOP_INDENT + UBound(data), L_COL))
    rng = data
    rng.Borders.LineStyle = XlLineStyle.xlContinuous
    rng.RowHeight = 15
End With

With ws.Sort
    Application.StatusBar = "Сортировка"
    .SortFields.Add Key:=ws.Cells(TOP_INDENT + 1, COL_NAME + 1), Order:=xlAscending
'    .SortFields.Add Key:=ws.Cells(TOP_INDENT + 1, COL_DATE), Order:=xlDescending
    .SortFields.Add Key:=ws.Cells(TOP_INDENT + 1, COL_NAME), Order:=xlAscending
    .SetRange rng
    .header = xlYes
    .Apply
    Application.StatusBar = "Контроль значений"
    Call SameProductsSameColor(rng)
End With

With ws
    Set rng = .Range(.Cells(TOP_INDENT, 1), .Cells(TOP_INDENT + UBound(data), L_COL))
    rng.AutoFilter
End With
Call Screen.Events(True)
Application.StatusBar = ""
Exit Sub

EXIT_SUB:
    MsgBox ("Не найдены НТД в папке НТД для анализа")
End Sub


Private Function GetAllData()
'Const sPath As String = "E:\Test Folder\Projects\НТД\НТД для анализа\"
Dim sPath As String
Dim counter As Integer
Dim doc_counter As Integer
Dim wb_name As String

'sPath = ActiveWorkbook.Path & FOLDER_ANALISYS
sPath = Paths.NTDPath
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFolder = oFSO.GetFolder(sPath)


counter = 1
doc_counter = FileCounter(sPath)
For Each oFile In oFolder.Files
    wb_name = oFile.name
    If InStr(wb_name, "$") = 0 Then
        Application.StatusBar = Format(((counter / doc_counter) * 100), "#,#0") & "%..." & wb_name
        data = GetData(data, "\" & wb_name, sPath)
        counter = counter + 1
    Else: End If
Next oFile
GetAllData = data
End Function



Private Function GetData(data As Variant, wb As String, wb_path As String)
Dim sql_data As Variant
Dim sql_str As String
Dim start As Long
Dim without_rep As Integer
'Dim calc_date As Date
'Dim calc_efio As String
'Dim calc_proj As String
Dim operation_name As String

sql_str = "SELECT F1, F2, F3, F4, F5, F6, F7, F8, F9, F10, F11, F12, F13 FROM [НТД$]"

'WHERE NOT (" & _
          "F1 IS null AND F2 IS null AND F3 IS null AND " & _
          "F4 IS null AND F5 IS null AND F6 IS null AND " & _
          "F7 IS null AND F8 IS null)
sql_data = SQL.SqlSelect(sql_str, wb, wb_path, "no")

If IsEmpty(sql_data) Then
    GetData = data
    Exit Function
Else: End If

'If IsEmpty(sql_data(COL_MANM_CALC, ROW_START - 1)) Or IsNull(sql_data(COL_MANM_CALC, ROW_START - 1)) Then
'    without_rep = 1
'Else
'    without_rep = 0
'End If

'last_row = 1
'For row = UBound(sql_data, 2) To LBound(sql_data, 2) Step -1
'    If Not IsNull(sql_data(0, row)) Then
'        last_row = row
'        Exit For
'    Else: End If
'Next row


If IsEmpty(data) Then
    start = 0 - ROW_START
    ReDim data(COL_HIER To L_COL, _
               LBound(sql_data, 2) To UBound(sql_data, 2))
Else
    start = UBound(data, 2) - ROW_START
    ReDim Preserve data(LBound(data) To UBound(data), _
                        LBound(data, 2) To UBound(data, 2) + UBound(sql_data, 2))
End If

'lenght = 0
'lenght = InStr(1, wb, "_") - 1
'temp = Left(wb, lenght)
'calc_proj = Right(temp, Len(temp) - 1)
'
'lenght = InStr(lenght + 2, wb, "_") - 1
'temp = Left(wb, lenght)
'temp = Right(temp, Len(temp) - Len(calc_proj) - 2)
'calc_date = Format(temp, "DD.MM.YYYY")
'
'lenght = InStrRev(wb, ".") - 1
'temp = Left(wb, lenght)
'lenght = InStrRev(wb, "_")
'calc_efio = Right(temp, Len(temp) - lenght)

wb = Left(Right(wb, Len(wb) - 1), Len(wb) - 6)

For row = ROW_START To UBound(sql_data, 2)
    If Not IsNull(sql_data(COL_HIERARCHY_CALC, row)) And Not IsEmpty(sql_data(COL_HIERARCHY_CALC, row)) And sql_data(COL_HIERARCHY_CALC, row) <> "" Then
        
        data(COL_HIER, row + start) = CStr(sql_data(COL_HIER_CALC, row))
        data(COL_NAME, row + start) = sql_data(COL_NAME_CALC, row)
        data(COL_DENO, row + start) = sql_data(COL_DENO_CALC, row)
        
        data(COL_MSR, row + start) = sql_data(COL_MSR_CALC, row)
        data(COL_NUM, row + start) = ToNum(sql_data(COL_NUM_CALC, row))
        
        data(COL_DIS, row + start) = ToNum(sql_data(COL_DIS_CALC, row))
        data(COL_DEF, row + start) = ToNum(sql_data(COL_DEF_CALC, row))
        data(COL_REP, row + start) = ToNum(sql_data(COL_REP_CALC, row))
        data(COL_RPR, row + start) = ToNum(sql_data(COL_RPR_CALC, row))
        
'        If without_rep = 0 Then
'            data(COL_RPR, row + start) = ToNum(sql_data(COL_RPR_CALC, row))
'        Else
'            data(COL_RPR, row + start) = "удалена"
'        End If
        
        data(COL_ASL, row + start) = ToNum(sql_data(COL_ASL_CALC - without_rep, row))
        data(COL_TUN, row + start) = ToNum(sql_data(COL_TUN_CALC - without_rep, row))
        data(COL_MAN, row + start) = ToNum(sql_data(COL_MAN_CALC - without_rep, row))
'        data(COL_MANM, row + start) = ToNum(sql_data(COL_MANM_CALC - without_rep, row))
        data(COL_TYPE, row + start) = sql_data(COL_TYPE_CALC - without_rep, row)
        data(COL_PROD, row + start) = wb
        data(COL_LINK_H, row + start) = row
        data(COL_LINK, row + start) = ">>>"
'        If Not IsNull(sql_data(COL_NAME_CALC, row)) Then
'            data(COL_DENO, row + start) = WsSubs.FindDeno(CStr(sql_data(COL_NAME_CALC, row)))
'        Else: End If
        
        
'    If Not IsNull(sql_data(COL_NORM_CALC, row)) And Not IsEmpty(sql_data(COL_NORM_CALC, row)) And sql_data(COL_NORM_CALC, row) <> "" Then
'        If sql_data(COL_HIERARCHY_CALC, row) <> "" Then
'            calc_deno = sql_data(COL_DENO_CALC, row)
'            If Not IsNull(calc_deno) Then
'                calc_deno = Replace(calc_deno, " ", "")
'            End If
'            data(COL_LEVEL_CALC, row + start) = sql_data(COL_LEVEL_CALC, row)
'            data(COL_HIERARCHY_CALC, row + start) = sql_data(COL_HIERARCHY_CALC, row)
'            data(COL_NAME_CALC, row + start) = sql_data(COL_NAME_CALC, row)
'            data(COL_DENO_CALC, row + start) = calc_deno
'            data(COL_NORM_CALC, row + start) = sql_data(COL_NORM_CALC, row)
'            data(COL_DATE_CALC, row + start) = calc_date
'            data(COL_EFIO_CALC, row + start) = calc_efio
'            data(COL_PROJ_CALC, row + start) = calc_proj
'            data(COL_COMM_CALC, row + start) = row
'            If data(COL_DENO_CALC, row + start) = "УИЕС.461434.001" Then
'                data(COL_DENO_CALC, row + start) = "УИЕС.461434.001"
'            Else: End If
'        Else:
'            If Not IsNull(sql_data(COL_NAME_CALC, row)) Then
'                operation_name = sql_data(COL_NAME_CALC, row)
'                operation_name = GetCorrectOperation(operation_name)
'
'                If operation_name <> "" Then
'                    data(COL_NAME_CALC, row + start) = GetCorrectOperation(operation_name)
'                    data(COL_NORM_CALC, row + start) = sql_data(COL_NORM_CALC, row)
'                Else: End If
'            Else: End If
'        End If
    Else
        Exit For
    End If
Next row
ReDim Preserve data(LBound(data) To UBound(data), _
                    LBound(data, 2) To row + start)

GetData = data

End Function



Sub ClearWS(ws As Worksheet)
last_row = ws.UsedRange.Rows.Count
With ws
    .AutoFilter.ShowAllData
    .Sort.SortFields.Clear
    .Cells.EntireColumn.Hidden = False
    .Cells.UnMerge
    .Cells.ClearContents
    .Range(.Cells(TOP_INDENT + 1, 1), Cells(last_row, 1)).EntireRow.Delete
    .Cells.Interior.ColorIndex = 2
    .Cells.Borders.LineStyle = None
End With

End Sub


Private Sub SetHeaderData(ws As Worksheet)
Dim rng As Range

With ws
    
    .Cells(2, COL_HIER) = "Индекс"
    .Cells(2, COL_NAME) = "Наименование"
    .Cells(2, COL_DENO) = "Децимальный" & Chr(10) & "номер"
    .Cells(2, COL_NUM) = "Кол-во"
    .Cells(2, COL_MSR) = "Ед. изм."
    .Cells(2, COL_DEF) = "Дефектация"
    .Cells(1, COL_DIS) = "Замена"
    .Cells(2, COL_DIS) = "Разборка"
    .Cells(2, COL_ASL) = "Сборка"
    .Cells(1, COL_REP) = "Ремонт" & Chr(10) & "на территории"
    .Cells(2, COL_REP) = "Заказчика"
    .Cells(2, COL_RPR) = "Исполнителя"
    .Cells(2, COL_TUN) = "Настройка"
    .Cells(2, COL_MAN) = "Изготовление"
    .Cells(1, COL_TYPE) = "Тип"
    .Cells(1, COL_PROD) = "НТД"
    .Cells(1, COL_LINK) = "Ссылка"
    .Cells(1, COL_TIME) = "Изготовление (Р)" ' & Chr(10) & "(расшифровка)"
    
    .Cells(2, COL_DEF).Orientation = 90

    .Cells(2, COL_DIS).Orientation = 90
    .Cells(2, COL_ASL).Orientation = 90

    .Cells(2, COL_REP).Orientation = 90
    .Cells(2, COL_RPR).Orientation = 90
    .Cells(1, COL_NUM).Orientation = 90
    .Cells(1, COL_MSR).Orientation = 90
    .Cells(1, COL_TUN).Orientation = 90
    .Cells(1, COL_MAN).Orientation = 90
    .Cells(1, COL_TIME).Orientation = 90
    .Cells(1, COL_LINK).Orientation = 90
    
    .Range(.Cells(1, COL_HIER), .Cells(2, COL_HIER)).Merge
    .Range(.Cells(1, COL_NAME), .Cells(2, COL_NAME)).Merge
    .Range(.Cells(1, COL_DENO), .Cells(2, COL_DENO)).Merge
    .Range(.Cells(1, COL_TYPE), .Cells(2, COL_TYPE)).Merge
    .Range(.Cells(1, COL_PROD), .Cells(2, COL_PROD)).Merge
    .Range(.Cells(1, COL_LINK), .Cells(2, COL_LINK)).Merge
    .Range(.Cells(1, COL_TIME), .Cells(2, COL_TIME)).Merge
    .Range(.Cells(1, COL_NUM), .Cells(2, COL_NUM)).Merge
    .Range(.Cells(1, COL_MSR), .Cells(2, COL_MSR)).Merge
    .Range(.Cells(1, COL_DEF), .Cells(2, COL_DEF)).Merge
    .Range(.Cells(1, COL_TUN), .Cells(2, COL_TUN)).Merge
    .Range(.Cells(1, COL_MAN), .Cells(2, COL_MAN)).Merge
    .Range(.Cells(1, COL_DIS), .Cells(1, COL_ASL)).Merge
    .Range(.Cells(1, COL_REP), .Cells(1, COL_RPR)).Merge
    
    Set rng = .Range(.Cells(1, 1), .Cells(1, L_COL + 1))
    rng.RowHeight = 30
    
    Set rng = .Range(.Cells(2, 1), .Cells(2, L_COL + 1))
    rng.RowHeight = 80
    
    Set rng = .Range(.Cells(1, 1), .Cells(TOP_INDENT, L_COL))
    rng.Borders.LineStyle = XlLineStyle.xlContinuous
    
    .Cells(1, COL_LINK_H).EntireColumn.Hidden = True
    
End With

End Sub


Private Function ToNum(value As Variant)

If IsNumeric(value) Then
    ToNum = CDec(value)
Else
    ToNum = value
End If

End Function


Private Sub SameProductsSameColor(rng As Range)
Const color1 As Integer = 19
Const color2 As Integer = 2
Const color_warning As Integer = 3
Dim color As Integer
Dim deno As String
Dim deno_next As String
Dim row As Long
Dim norm
Dim norm_next



color = color1
criteria = GetCriteria(rng, 1)
'deno = rng.Rows(1).Cells(1, COL_DENO).value
rng.Rows(1).Interior.ColorIndex = color


For row = 2 To rng.Rows.Count
    
    criteria_next = GetCriteria(rng, row)
'    deno_next = rng.Rows(row).Cells(1, COL_DENO).value

    If criteria <> criteria_next Then
        
        If color = color1 Then
            color = color2
        Else
            color = color1
        End If
        rng.Rows(row).Interior.ColorIndex = color
    Else
        rng.Rows(row).Interior.ColorIndex = color
        
        For col = COL_MSR To COL_TYPE
            norm = rng.Rows(row - 1).Cells(1, col).Value2
            norm_next = rng.Rows(row).Cells(1, col).Value2
            If norm_next <> norm Then
                rng.Rows(row - 1).Cells(1, col).Interior.ColorIndex = color_warning
                rng.Rows(row).Cells(1, col).Interior.ColorIndex = color_warning
                rng.Rows(row - 1).Cells(1, COL_NAME).Interior.ColorIndex = color_warning
                rng.Rows(row).Cells(1, COL_NAME).Interior.ColorIndex = color_warning
            Else: End If
        Next col
    End If
    criteria = criteria_next
Next row
End Sub

Private Function GetCriteria(rng As Range, row As Long)
Dim criteria As String

criteria = rng.Rows(1).Cells(row, COL_DENO).value
If criteria = "" Then
    criteria = rng.Rows(1).Cells(row, COL_NAME).value
Else: End If
GetCriteria = criteria
End Function

Sub CalculationLink(row As Long, col As Integer)
Dim wb_name As String
Dim row_num As Integer


row_num = ThisWorkbook.Worksheets("Таблица").Cells(row, COL_LINK_H) + 1
wb_name = ThisWorkbook.Worksheets("Таблица").Cells(row, COL_PROD)

On Error GoTo EXIT_SUB

Paths.NTDPath
ActiveWorkbook.FollowHyperlink (Paths.NTDPath & "/" & wb_name & ".xlsm")

'ActiveWorkbook.FollowHyperlink (ActiveWorkbook.Path & FOLDER_ANALISYS & wb_name & ".xlsx")
ActiveWorkbook.Worksheets("НТД").Activate
ActiveWorkbook.Worksheets("НТД").Range("A" & row_num).Select
Exit Sub

EXIT_SUB:
    MsgBox ("Не найден файл " & wb_name & Chr(10) & "в папке " & ActiveWorkbook.Path & FOLDER_ANALISYS)
End Sub
