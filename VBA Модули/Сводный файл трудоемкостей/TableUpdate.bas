Attribute VB_Name = "TableUpdate"
Sub main()
'Dim ws As Worksheet

'ws = ActiveSheet()

Select Case ActiveSheet().name

Case "�������"
    Call TimeCollector.main
Case "������ �� ��"
    Call DataMK.main
Case "������ �� �� (�� pdf)"
    Call DataMKupd.main
Case Else
End Select

End Sub
