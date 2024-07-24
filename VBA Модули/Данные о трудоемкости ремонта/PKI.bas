Attribute VB_Name = "PKI"
Const col_name As Integer = 1
Const col_sett As Integer = col_name + 1
Const col_kfft As Integer = col_sett + 1
Const col_trud As Integer = col_kfft + 1
Const col_last As Integer = col_trud


Function PKI(data As Variant)
Dim name As String

data_pki = GetPKI()
For row = 3 To UBound(data)
    If data(row, Calculation.col_type) = "ПКИ" Then 'Or data(row, Calculation.col_type) = "Деталь"
        name = LCase(data(row, Calculation.col_name))
        name = Replace(name, " ", "", 1)
        For i = 0 To 9
            name = Replace(name, i, "", 1)
        Next i
        For row_pki = LBound(data_pki) + 1 To UBound(data_pki)
            If InStr(1, name, LCase(data_pki(row_pki, col_name)), vbTextCompare) <> 0 Then
                data(row, Calculation.col_def_one_calc) = data_pki(row_pki, col_trud) '* Format(0.01 * WorksheetFunction.RandBetween(70, 130), "#,#0.0")
                Exit For
            End If
        Next row_pki
    End If
Next row
PKI = data
End Function


Private Function GetPKI()
Dim ws_pki As Worksheet

Set ws_pki = ThisWorkbook.Worksheets("ПКИ (оценка)")
l_row = DocumentAttribute.LastRow(ws_pki, col_name) + 1
With ws_pki
    data = .Range(.Cells(1, col_name), .Cells(l_row, col_last))
End With
GetPKI = data
End Function



