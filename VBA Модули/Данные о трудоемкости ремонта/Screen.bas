Attribute VB_Name = "Screen"
'Включение/Отключение реакции на события и пересчет листа
Sub Events(events_on As Boolean)
Application.EnableEvents = events_on
Application.ScreenUpdating = events_on
If events_on Then
    Application.Calculation = xlAutomatic
Else
    Application.Calculation = xlManual
End If
End Sub

Sub SaveAutoFilter(w As Worksheet, currentFiltRange As String, filterArray As Variant)
With w.AutoFilter
    currentFiltRange = .Range.Address
    With .Filters
        ReDim filterArray(1 To .Count, 1 To 3)
        For f = 1 To .Count
            With .item(f)
                If .On Then
                    filterArray(f, 1) = .Criteria1
                    If .Operator Then
                        filterArray(f, 2) = .Operator
                        filterArray(f, 3) = .Criteria2 'simply delete this line to make it work in Excel 2010
                    End If
                End If
            End With
        Next f
    End With
End With
w.AutoFilter.ShowAllData
End Sub

Sub RestoreAutoFilter(w As Worksheet, currentFiltRange As String, filterArray As Variant)
Dim col As Integer

    For col = 1 To UBound(filterArray, 1)
        If Not IsEmpty(filterArray(col, 1)) Then
            If filterArray(col, 2) Then
                On Error Resume Next
                w.Range(currentFiltRange).AutoFilter field:=col, Criteria1:=filterArray(col, 1), Operator:=filterArray(col, 2), Criteria2:=filterArray(col, 3)
            
            Else
                w.Range(currentFiltRange).AutoFilter field:=col, Criteria1:=filterArray(col, 1)
                On Error Resume Next
            End If
        End If
    Next col
End Sub


