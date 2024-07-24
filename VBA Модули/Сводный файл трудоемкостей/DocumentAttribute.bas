Attribute VB_Name = "DocumentAttribute"
'Код вида ТД
Function TdTypeDoc(td_denotation_number As String)
Select Case Right(Left(td_denotation_number, 7), 2)
    Case Is = "25"
        TdTypeDoc = "ТИ"
    Case Is = "10"
        TdTypeDoc = "МК"
    Case Is = "50"
        TdTypeDoc = "КТП"
    Case Is = "55"
        TdTypeDoc = "КТ(Г)ТП"
    Case Is = "20"
        TdTypeDoc = "КЭ"
    Case Is = "60"
        TdTypeDoc = "ОК"
    Case Is = "57"
        TdTypeDoc = "КТО"
    Case Is = "75"
        TdTypeDoc = "ТНК"
    Case Is = "30"
        TdTypeDoc = "КК"
    Case Is = "62"
        TdTypeDoc = "КН"
    Case Is = "59"
        TdTypeDoc = "КТИ"
    Case Is = "67"
        TdTypeDoc = "ККИ"
    Case Is = "66"
        TdTypeDoc = "КРИ"
    Case Is = "01"
        TdTypeDoc = "ТЛ"
    Case Is = "02"
        TdTypeDoc = "ТЛ(Д)"
    Case Is = "04"
        TdTypeDoc = "ТЛ(ДВр)"
    Case Is = "05"
        TdTypeDoc = "ТЛ(Пр)"
    Case Is = "06"
        TdTypeDoc = "ТЛ(Д)"
    Case Is = "07"
        TdTypeDoc = "ТЛ(ДИ)"
    Case Is = "09"
        TdTypeDoc = "ТЛ(ДСт)"
    Case Is = "80"
        TdTypeDoc = "ВДП"
    Case Is = "40"
        TdTypeDoc = "ВТД"
    Case Is = "41"
        TdTypeDoc = "ВО"
    Case Is = "72"
        TdTypeDoc = "ВСИ"
    Case Is = "45"
        TdTypeDoc = "ВСИ"
    Case Is = "71"
        TdTypeDoc = "ВП"
    Case Is = "00"
        TdTypeDoc = "ВНТ"
    Case Is = "43"
        TdTypeDoc = "ВМ"
    Case Is = "48"
        TdTypeDoc = "ВУН"
    Case Is = "47"
        TdTypeDoc = "ВСН"
    Case Is = "46"
        TdTypeDoc = "ВОБ"
    Case Is = "42"
        TdTypeDoc = "ВО"
    Case Is = "44"
        TdTypeDoc = "ВТП"
    Case Is = "78"
        TdTypeDoc = "ВД"
    Case Is = "77"
        TdTypeDoc = "ВДО"
    Case Is = "79"
        TdTypeDoc = "ВСТ"
    Case Is = "70"
        TdTypeDoc = "ТВ"
    Case Else
End Select
End Function


'Код вида ТД по организации
Function TdTypeOrg(td_denotation_number As String)
Select Case Right(Left(td_denotation_number, 8), 1)
     Case Is = "1"
     TdTypeOrg = "Е"
     Case Is = "2"
     TdTypeOrg = "Т"
     Case Is = "3"
     TdTypeOrg = "Г"
     Case Is = "0"
     TdTypeOrg = "Н/У"
     Case Else
End Select
End Function


'Код вида ТД по методу
Function TdTypeMethod(td_denotation_number As String)
Select Case Right(Left(td_denotation_number, 10), 2)
     Case Is = "00"
     TdTypeMethod = "Без указания"
     Case Is = "88"
     TdTypeMethod = "Сборка"
     Case Is = "01"
     TdTypeMethod = "Общего назначения"
     Case Is = "85"
     TdTypeMethod = "Электромонтаж"
     Case Is = "80"
     TdTypeMethod = "Пайка"
     Case Is = "02"
     TdTypeMethod = "Технический контроль"
     Case Is = "21"
     TdTypeMethod = "Обработка давлением"
     Case Is = "41"
     TdTypeMethod = "Обработка резанием"
     Case Is = "06"
     TdTypeMethod = "Испытания"
     Case Is = "08"
     TdTypeMethod = "Консервация и упаковывание"
     Case Is = "73"
     TdTypeMethod = "Получение покрытий лакокрасочных (органических)"
     Case Is = "71"
     TdTypeMethod = "Получение покрытия (металлического и неметаллического неорганического)"
     Case Is = "75"
     TdTypeMethod = "Электрофизическая, электрохимическая и радиационная обработка"
     Case Is = "65"
     TdTypeMethod = "Порошковая металлургия"
     Case Is = "55"
     TdTypeMethod = "Фотохимико-физическая обработка"
     Case Is = "60"
     TdTypeMethod = "Формообразование из полимерных материалов, керамики, стекла и резины"
     Case Is = "50"
     TdTypeMethod = "Термообработка"
     Case Is = "10"
     TdTypeMethod = "Литье металлов и сплавов"
     Case Is = "04"
     TdTypeMethod = "Перемещение"
     Case Is = "90"
     TdTypeMethod = "Сварка"
     Case Else
End Select
End Function

'Поиск номера последнего ряда
Function LastRow(ws As Worksheet, column As Integer)
Dim l_row As Long
l_row = ws.Cells(1048576, column).End(xlUp).row
LastRow = l_row
End Function

'Код по названию документа
Function TdTypeDocCode(doc_name As String)
Dim code As String

Select Case doc_name
Case Is = "Технологическая инструкция"
    code = "25"
Case Is = "Маршрутная карта"
    code = "10"
Case Is = "Карта технологического процесса"
    code = "50"
Case Is = "Карта типового (группового) технологического процесса"
    code = "55"
Case Is = "Карта эскизов"
    code = "20"
Case Is = "Операционная карта"
    code = "60"
Case Is = "Карта типовой (групповой) операции"
    code = "57"
Case Is = "Технико-нормировочная карта"
    code = "75"
Case Is = "Комплектовочная карта"
    code = "30"
Case Is = "Карта наладки"
    code = "62"
Case Is = "Карта технологической информации"
    code = "59"
Case Is = "Карта кодирования информации"
    code = "67"
Case Is = "Карта расчета информации"
    code = "66"
Case Is = "Комплект технологической документации"
    code = "01"
Case Is = "Комплект документов ТП (операции)"
    code = "02"
Case Is = "Комплект временных документов ТП (операции)"
    code = "04"
Case Is = "Комплект проектной технологической документации"
    code = "05"
Case Is = "Комплект директивной технологической документации"
    code = "06"
Case Is = "Комплект документов ТП (операции) информационного назначения"
    code = "07"
Case Is = "Стандартный комплект документов ТП (операции)"
    code = "09"
Case Is = "Ведомость держателей подлинников"
    code = "80"
Case Is = "Ведомость технологических документов"
    code = "40"
Case Is = "Ведомость технологических маршрутов"
    code = "41"
Case Is = "Ведомость операций"
    code = "72"
Case Is = "Ведомость сборки изделия"
    code = "45"
Case Is = "Ведомость применяемости"
    code = "71"
Case Is = "Ведомость нормирования труда"
    code = "00"
Case Is = "Ведомость материалов"
    code = "43"
Case Is = "Ведомость удельных норм расхода материалов"
    code = "48"
Case Is = "Ведомость специфицированных норм расхода материалов"
    code = "47"
Case Is = "Ведомость оборудования"
    code = "46"
Case Is = "Ведомость оснастки"
    code = "42"
Case Is = "Ведомость ДСЕ к типовому (групповому) ТП (операции)"
    code = "44"
Case Is = "Ведомость дефектации"
    code = "78"
Case Is = "Ведомость деталей, изготовленных из отходов"
    code = "77"
Case Is = "Ведомость стержней"
    code = "79"
Case Is = "Технологическая ведомость"
    code = "70"
Case Else
    code = ""
End Select

TdTypeDocCode = code
End Function


'Сокращение названия документа
Function TdTypeDocAbbrev(doc_name As String)
Dim code As String

Select Case doc_name
Case Is = "Технологическая инструкция"
    code = "ТИ"
Case Is = "Маршрутная карта"
    code = "МК"
Case Is = "Карта технологического процесса"
    code = "КТП"
Case Is = "Карта типового (группового) технологического процесса"
    code = "КТ(Г)ТП"
Case Is = "Карта эскизов"
    code = "КЭ"
Case Is = "Операционная карта"
    code = "ОК"
Case Is = "Карта типовой (групповой) операции"
    code = "КТО"
Case Is = "Технико-нормировочная карта"
    code = "ТНК"
Case Is = "Комплектовочная карта"
    code = "КК"
Case Is = "Карта наладки"
    code = "КН"
Case Is = "Карта технологической информации"
    code = "КТИ"
Case Is = "Карта кодирования информации"
    code = "ККИ"
Case Is = "Карта расчета информации"
    code = "КРИ"
Case Is = "Комплект технологической документации"
    code = "ТЛ"
Case Is = "Комплект документов ТП (операции)"
    code = "ТЛ(Д)"
Case Is = "Комплект временных документов ТП (операции)"
    code = "ТЛ(ДВр)"
Case Is = "Комплект проектной технологической документации"
    code = "ТЛ(Пр)"
Case Is = "Комплект директивной технологической документации"
    code = "ТЛ(Д)"
Case Is = "Комплект документов ТП (операции) информационного назначения"
    code = "ТЛ(ДИ)"
Case Is = "Стандартный комплект документов ТП (операции)"
    code = "ТЛ(ДСт)"
Case Is = "Ведомость держателей подлинников"
    code = "ВДП"
Case Is = "Ведомость технологических документов"
    code = "ВТД"
Case Is = "Ведомость технологических маршрутов"
    code = "ВО"
Case Is = "Ведомость операций"
    code = "ВСИ"
Case Is = "Ведомость сборки изделия"
    code = "ВСИ"
Case Is = "Ведомость применяемости"
    code = "ВП"
Case Is = "Ведомость нормирования труда"
    code = "ВНТ"
Case Is = "Ведомость материалов"
    code = "ВМ"
Case Is = "Ведомость удельных норм расхода материалов"
    code = "ВУН"
Case Is = "Ведомость специфицированных норм расхода материалов"
    code = "ВСН"
Case Is = "Ведомость оборудования"
    code = "ВОБ"
Case Is = "Ведомость оснастки"
    code = "ВО"
Case Is = "Ведомость ДСЕ к типовому (групповому) ТП (операции)"
    code = "ВТП"
Case Is = "Ведомость дефектации"
    code = "ВД"
Case Is = "Ведомость деталей, изготовленных из отходов"
    code = "ВДО"
Case Is = "Ведомость стержней"
    code = "ВСТ"
Case Is = "Технологическая ведомость"
    code = "ТВ"
Case Else
    code = ""
End Select

TdTypeDocAbbrev = code
End Function


'Код по типу организации
Function TdTypeOrgCode(doc_name As String)
Dim code As String

Select Case UserFormBase.VidTDOrg.value
Case Is = "Единичный процесс (операция)"
    code = "1"
Case Is = "Группой процесс (операция)"
    code = "2"
Case Is = "Типовой процесс (операция)"
    code = "3"
Case Is = "Без указания"
    code = "0"
Case Else
    code = ""
End Select

TdTypeOrgCode = code
End Function


'Сокращение по типу организации
Function TdTypeOrgAbbrev(doc_name As String)
Dim code As String

Select Case UserFormBase.VidTDOrg.value
Case Is = "Единичный процесс (операция)"
    code = "Е"
Case Is = "Группой процесс (операция)"
    code = "Т"
Case Is = "Типовой процесс (операция)"
    code = "Г"
Case Is = "Без указания"
    code = "Н/У"
Case Else
    code = ""
End Select

TdTypeOrgAbbrev = code
End Function

Function TdTypeMethodCode(doc_name As String)
Dim code As String

Select Case doc_name
Case Is = "Без указания"
    code = "00"
Case Is = "Сборка"
    code = "88"
Case Is = "Общего назначения"
    code = "01"
Case Is = "Электромонтаж"
    code = "85"
Case Is = "Пайка"
    code = "80"
Case Is = "Технический контроль"
    code = "02"
Case Is = "Обработка давлением"
    code = "21"
Case Is = "Обработка резанием"
    code = "41"
Case Is = "Испытания"
    code = "06"
Case Is = "Консервация и упаковывание"
    code = "08"
Case Is = "Получение покрытий лакокрасочных (органических)"
    code = "73"
Case Is = "Получение покрытия (металлического и неметаллического неорганического)"
    code = "71"
Case Is = "Электрофизическая, электрохимическая и радиационная обработка"
    code = "75"
Case Is = "Порошковая металлургия"
    code = "65"
Case Is = "Фотохимико-физическая обработка"
    code = "55"
Case Is = "Формообразование из полимерных материалов, керамики, стекла и резины"
    code = "60"
Case Is = "Термообработка"
    code = "50"
Case Is = "Литье металлов и сплавов"
    code = "10"
Case Is = "Перемещение"
    code = "04"
Case Is = "Сварка"
    code = "90"
Case Else
    code = ""
End Select

TdTypeMethodCode = code
End Function


Function InArray(arr As Variant, value As String)
InArray = False
For Each arr_val In arr
    If value = arr_val Then
        InArray = True
        Exit For
    Else: End If
Next arr_val
End Function

Sub AddNewWS(name As String)
Dim ws As Worksheet
Dim exist As Boolean

exist = False
exist = WorksheetExists(name)

If exist Then
    Set ws = ThisWorkbook.Worksheets(name)
Else
    With ThisWorkbook
        Set ws = .sheets.Add(After:=.sheets(.sheets.Count))
        ws.name = name
    End With
End If

End Sub


Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
Dim sht As Worksheet

If wb Is Nothing Then Set wb = ThisWorkbook
On Error Resume Next
Set sht = wb.sheets(shtName)
On Error GoTo 0
WorksheetExists = Not sht Is Nothing
End Function


Function Reverse2dArray(data As Variant)
Dim data_new As Variant
Dim row_new As Long
data_new = data
row_new = LBound(data)
For row = UBound(data) To 1 Step -1
    For col = LBound(data, 2) To UBound(data, 2)
        data_new(row_new, col) = data(row, col)
    Next col
    row_new = row_new + 1
Next row
Reverse2dArray = data_new
End Function


Function Transpose2dArray(arr As Variant)
ReDim new_arr(LBound(arr, 2) To UBound(arr, 2), LBound(arr) To UBound(arr))
For row = LBound(new_arr) To UBound(new_arr)
    For col = LBound(new_arr, 2) To UBound(new_arr, 2)
        new_arr(row, col) = arr(col, row)
    Next col
Next row
Transpose2dArray = new_arr
End Function


Function FileCounter(sPath As String)
Dim doc_counter As Integer

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFolder = oFSO.GetFolder(sPath)
 
doc_counter = 1
'Цикл определения количества файлов в папке
For Each oFile In oFolder.Files
    Application.StatusBar = "Количество документов = " & doc_counter
    wb_name = oFile.name
    doc_counter = doc_counter + 1
Next
        
FileCounter = doc_counter
End Function
