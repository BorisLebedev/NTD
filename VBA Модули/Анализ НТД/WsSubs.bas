Attribute VB_Name = "WsSubs"
Sub WsSelectionChanged(row As Long, col As Long)

If row > 1 Then
    Select Case col
        Case 3
            Call ValidationOperation(row, col)
        End Select
Else: End If

End Sub


Sub ValidationOperation(row As Long, col As Long)

Cells(row, col).EntireColumn.Validation.Delete
If Cells(row, Calculation.col_hierarchy) = "" Then
    With Cells(row, Calculation.col_operation).Validation
        .Delete
        .Add Type:=xlValidateList, Operator:=xlBetween, Formula1:="=OPERATIONS"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = False
    End With
Else: End If
End Sub


Sub WsChanged(Target As Range, row As Long, col As Long)

If Target.row > 1 Then
    Select Case Target.column
        Case 1
            Call SetIndex(Target)
        Case 3
            Call ValidationOperationChanged(Target)
        End Select
Else: End If
End Sub


Private Sub ValidationOperationChanged(Target As Range)
Dim deno As String
Dim text As String

If IsEmpty(Calculation.OPERATIONS) Then
    Calculation.OPERATIONS = GetOperationArray()
Else: End If
On Error GoTo ExitCurrentSub
If Target.column = Calculation.COL_NAME And Cells(Target.row, Calculation.col_hierarchy) = "" Then
    For row = LBound(Calculation.OPERATIONS) To UBound(Calculation.OPERATIONS)
        If Target.value = Calculation.OPERATIONS(row, 1) Then
            Cells(Target.row, Calculation.COL_DENO) = Calculation.OPERATIONS(row, 2)
            Exit For
        Else: End If
    Next row
Else
    If Cells(Target.row, Calculation.COL_DENO) = "" Then
        text = Cells(Target.row, Calculation.COL_NAME)
        deno = FindDeno(text)
        If deno <> "" Then
            Cells(Target.row, Calculation.COL_NAME) = Replace(text, " " & deno, "")
            Cells(Target.row, Calculation.COL_DENO) = deno
        Else: End If
    Else: End If
End If
ExitCurrentSub:

End Sub


Private Sub SetIndex(Target As Range)
Dim row As Long
Dim level As Integer
Dim index As String
Dim rng As Range
Dim ws As Worksheet
Dim data As Variant

Call Screen.Events(False)
Set ws = ActiveSheet()
l_row = DocumentAttribute.LastRow(ws, 1)
With ws
    If l_row > Calculation.TOP_INDENT Then
        Set rng = .Range(.Cells(Calculation.TOP_INDENT + 1, 1), .Cells(l_row, Calculation.col_hierarchy))
        data = rng.value
        data = CalcIndex(data)
        rng = data
    Else: End If
End With
Call Screen.Events(True)
End Sub


Private Function CalcIndex(data As Variant)
Dim level As Integer
Dim sub_level As Integer
Dim top_level As Integer
Dim counter As Integer
Dim sub_row As Long
Dim index As String
Dim top_index As String

For row = LBound(data) To UBound(data)
    If data(row, 1) = 0 And Not IsEmpty(data(row, 1)) Then
        index = "Изделие"
    Else
        If IsEmpty(data(row, 1)) Or data(row, 1) = "" Then
            index = ""
        Else
            counter = 1
            sub_row = row
            level = data(row, 1)
            top_level = level - 1
            sub_level = level + 1
            Do While sub_level >= level And sub_row <> LBound(data)
                sub_row = sub_row - 1
                If Not IsEmpty(data(sub_row, Calculation.col_level)) And data(sub_row, Calculation.col_level) <> "" Then
                    sub_level = data(sub_row, Calculation.col_level)
                    If sub_level = level Then
                        counter = counter + 1
                    Else: End If
                Else: End If
            Loop
            top_index = data(sub_row, Calculation.col_hierarchy)
            
            If top_index = "Изделие" Then
                index = CStr("" & counter)
            Else
                index = CStr(top_index & "." & counter)
            End If
            
        End If
    End If
    data(row, 2) = index
Next row

CalcIndex = data

End Function


Private Function GetIndexFromLevel(row As Long, level As Integer)
Dim index As String

main_level = level
Do While Not (level = main_level Or row = 0)
    row = row - 1
    level = Cells(row, 1)
Loop
index = Cells(row - 1, 2)

If index = "Изделие" Then
    index = 1
Else
    point_pos = InStrRev(index, ".")
    index_num = CInt(Right(index, Len(index) - point_pos)) + 1
    index = index_num
End If


GetIndexFromLevel = index
End Function


'Найти все децимальные номера в тексте
Function FindDeno(perehod As String, Optional only_base As Boolean = False)
Dim reg_ex As New RegExp
Dim mc As MatchCollection
Dim item As Variant
Dim dec_dict As Object
Dim dec_num As String
Dim iot_str As String
Dim arrayList As Object

reg_ex.Global = True
If only_base Then
    reg_ex.pattern = "[А-Я][А-Я][А-Я][А-Я]\.[0-9]{6}\.[0-9]{3}"
Else
    reg_ex.pattern = "[А-Я][А-Я][А-Я][А-Я]\.(([0-9]{6}\.([0-9]{3}\-[0-9]{2}|[0-9]{3})([А-Я][0-9][0-9]|[А-Я][А-Я][0-9]|[А-Я][А-Я]|[А-Я][0-9]|))|[0-9]{5}.[0-9]{5}|[0-9]{5}\-[0-9]{2})"
End If


Set mc = reg_ex.Execute(perehod)

Set dec_dict = CreateObject("Scripting.Dictionary")
For Each item In mc:
    dec_dict(item.value) = 1
Next item

iot_str = ""
For Each item In dec_dict:
    iot_str = iot_str & ", " & item
Next item
If Len(iot_str) <> 0 Then
    FindDeno = Right(iot_str, (Len(iot_str) - 2))
Else
    FindDeno = ""
End If
End Function


'Найти все децимальные номера в тексте
Function FindByPattern(perehod As String, pattern As String)
Dim reg_ex As New RegExp
Dim mc As MatchCollection
Dim item As Variant
Dim dec_dict As Object
Dim dec_num As String
Dim iot_str As String
Dim arrayList As Object

reg_ex.Global = True
reg_ex.pattern = pattern

Set mc = reg_ex.Execute(perehod)

Set dec_dict = CreateObject("Scripting.Dictionary")
For Each item In mc:
    dec_dict(item.value) = 1
Next item

iot_str = ""
For Each item In dec_dict:
    iot_str = iot_str & ", " & item
Next item
If Len(iot_str) <> 0 Then
    FindByPattern = Right(iot_str, (Len(iot_str) - 2))
Else
    FindByPattern = ""
End If
End Function

