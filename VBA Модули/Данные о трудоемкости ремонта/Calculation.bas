Attribute VB_Name = "Calculation"
Global Const col_level As Integer = 1
Global Const col_hierarchy As Integer = col_level + 1
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

Global Const col_def_one_calc As Integer = col_new_one + 1
Global Const col_dis_one_calc As Integer = col_def_one_calc + 1
Global Const col_ass_one_calc As Integer = col_dis_one_calc + 1
Global Const col_rpr_one_calc As Integer = col_ass_one_calc + 1
Global Const col_rpl_one_calc As Integer = col_rpr_one_calc + 1
Global Const col_tun_one_calc As Integer = col_rpl_one_calc + 1
Global Const col_new_one_calc As Integer = col_tun_one_calc + 1

Global Const col_type As Integer = col_new_one_calc + 1

Global Const l_col As Integer = col_type
Global Const top_indent As Integer = 3


Sub Main()
Const index_col As Integer = col_name
Dim data As Variant
Dim ws As Worksheet
Dim l_row As Long
Dim rng As Range
Dim currentFiltRange As String
Dim filterArray As Variant

Call Screen.Events(False)
Call ShowAllSheets(True)
Set ws = ThisWorkbook.Worksheets("Расчет")
With ws
    Call Screen.SaveAutoFilter(ws, currentFiltRange, filterArray)
    .AutoFilter.ShowAllData

    l_row = DocumentAttribute.LastRow(ws, col_name) + 1
    If l_row > top_indent Then
        Set rng = .Range(.Cells(top_indent, 1), .Cells(l_row - 1, l_col))
    Else
        Set rng = .Range(.Cells(top_indent, 1), .Cells(top_indent + 1, l_col))
    End If
    
    data = rng.value
    data = GetDeno(data)
    data = GetMsr(data)
    data = GetTypesOfProduct(data)
    data = Based.GetBaseData(data)
    data = PKI.PKI(data)
    rng = data
    Call Screen.RestoreAutoFilter(ActiveSheet, currentFiltRange, filterArray)
    'Call SetFormatOfData(rng, data)
End With
Call ShowAllSheets(False)
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


Function GetDeno(data As Variant)
Dim text As String
Dim row As Long

For row = 2 To UBound(data)
    If IsEmpty(data(row, col_deno)) Then
        text = data(row, col_name)
        data(row, col_deno) = WsSubs.FindDeno(text)
        data(row, col_name) = RTrim(Replace(data(row, col_name), data(row, col_deno), ""))
    End If
Next row

GetDeno = data
End Function


Function GetMsr(data As Variant)
Dim num As String
Dim row As Long

For row = 2 To UBound(data)
    If IsEmpty(data(row, col_msr)) And Not IsEmpty(data(row, col_num)) Then
        num = data(row, col_num)
        If CInt(num) = CDbl(num) Then
            data(row, col_msr) = "шт"
        End If
    End If
Next row

GetMsr = data
End Function


Function GetTypesOfProduct(data As Variant)
Dim text As String
Dim name As String
Dim row As Long


data(2, col_type) = "Изделие"

For row = 3 To UBound(data)
    If IsEmpty(data(row, col_type)) Then
        
        'ПКИ
        If IsEmpty(data(row, col_deno)) Or data(row, col_deno) = "" Then
            data(row, col_type) = "ПКИ"
        Else
            text = data(row, col_deno)
            name = data(row, col_name)
            
            'Комплект
            If WsSubs.FindByPattern(name, "([К][М][Ч])|([Кк][о][м][п][л][е][к][т])") <> "" Then
                data(row, col_type) = "Комплект"
            Else
                
                'ЗИП
                If WsSubs.FindByPattern(name, "([З][И][П])") <> "" Then
                    data(row, col_type) = "ЗИП"
                Else
                    
                    'Деталь
                    If WsSubs.FindByPattern(text, "[А-Я]{4}\.[7][0-8]") <> "" Then
                        data(row, col_type) = "Деталь (ПКИ)"
                    Else
                        
                        'Детали / Кабели
                        If WsSubs.FindByPattern(text, "[А-Я]{4}\.[6][0-9]") <> "" Then
                            If WsSubs.FindByPattern(name, "([Ш][и][н][а])|([ ][ш][и][н][а])") <> "" Then
                                data(row, col_type) = "Деталь (ПКИ)"
                            Else
                            
                                'Кабель
                                If WsSubs.FindByPattern(name, "([Кк][а][б][е][л])|([Жж][г][у][т])") <> "" Then
                                    data(row, col_type) = "Кабель"
                                End If
                            End If
                        Else
                        
                            'СПО
                            If WsSubs.FindByPattern(text, "[А-Я]{4}\.[0-9]{5}\-") <> "" Then
                                data(row, col_type) = "СПО"
                            Else
                            
                            End If
                    
                        End If
                        
                    End If
                End If
            End If
        End If
    End If
Next row

GetTypesOfProduct = data
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


Function HasNextLevel(data As Variant, row As Long)

level_current = data(row, Calculation.col_level)
If row = UBound(data) Then
    level_next = level_current
Else:
    level_next = data(row + 1, Calculation.col_level)
End If
HasNextLevel = level_current < level_next

End Function


Sub ShowAllSheets(show As Boolean)

For Each ws In ThisWorkbook.Sheets
    If Not (ws.name = "НТД" Or ws.name = "Расчет" Or ws.name = "Типы") Then
        ws.Visible = show
    Else: End If
Next ws
End Sub
