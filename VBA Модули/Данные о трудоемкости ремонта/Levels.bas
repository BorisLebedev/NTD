Attribute VB_Name = "Levels"
Sub LevelsByIndex()
Const index_col As Integer = 7
Dim data As Variant
Dim ws As Worksheet
Dim l_row As Long
Dim rng As Range
Dim rng_lvl As Range

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
    data = GetLevels(data)
    
    data_lvl = GetSubArray(data, Calculation.col_level)
    Set rng_lvl = .Range(.Cells(Calculation.top_indent + 1, Calculation.col_level), .Cells(l_row, Calculation.col_level))
    rng_lvl = data_lvl

    Call level(l_row)
End With
Call Screen.Events(True)
End Sub


Private Function GetLevels(data As Variant)
Dim row As Long
Dim index_hierarchy As String
Dim level As Integer

l_row = UBound(data)
For row = UBound(data) To 2 Step -1
    data(row, Calculation.col_level) = ""
    index_hierarchy = data(row, Calculation.col_hierarchy)
    If index_hierarchy <> "" Then
        data(row, Calculation.col_level) = Calculation.GetLevel(index_hierarchy)
    Else: End If
Next
GetLevels = data
End Function


Private Sub level(l_row As Long)
Dim level As Integer
Dim group As Integer

'LastSTR = Range(Cells(6, 11), Cells(100000, 11)).End(xlDown).row

On Error GoTo err
For i = 1 To 8
    Range(Cells(Calculation.top_indent + 1, 1), Cells(l_row, 1)).EntireRow.Ungroup
Next i
level = 0
err:
For row = Calculation.top_indent + 2 To l_row
    If Not IsEmpty(Cells(row, Calculation.col_level)) And Cells(row, Calculation.col_level) <> "" Then
        level = Cells(row, Calculation.col_level)
        group = level - 1
    Else
        group = level
    End If
    
    For i = 1 To group
        If i < 8 Then
            Cells(row, Calculation.col_level).EntireRow.group
        Else: End If
    Next i
    
    
Next row

End Sub


