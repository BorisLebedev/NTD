Attribute VB_Name = "������������_�����������"
'����� ��� ����������� ������ � ������
Function �����������(perehod As String)
Dim reg_ex As New RegExp
Dim mc As MatchCollection
Dim item As Variant
Dim dec_dict As Object
Dim dec_num As String
Dim iot_str As String
Dim arrayList As Object

reg_ex.Global = True
reg_ex.pattern = "[�-�][�-�][�-�][�-�]\.(([0-9]{6}\.([0-9]{3}\-[0-9]{2}|[0-9]{3})([�-�][0-9][0-9]|[�-�][�-�][0-9]|[�-�][�-�]|[�-�][0-9]|))|[0-9]{5}.[0-9]{5}|[0-9]{5}\-[0-9]{2})"
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
    ����������� = Right(iot_str, (Len(iot_str) - 2))
Else
    ����������� = ""
End If
End Function


'����� ��� ����������� ������ � ������
Function ������������(perehod As String)

deno = �����������(perehod)
������������ = Replace(perehod, deno, "")


End Function
