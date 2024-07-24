Attribute VB_Name = "TableUpdate"
Sub main()
'Dim ws As Worksheet

'ws = ActiveSheet()

Select Case ActiveSheet().name

Case "Таблица"
    Call TimeCollector.main
Case "Данные из МК"
    Call DataMK.main
Case "Данные из МК (из pdf)"
    Call DataMKupd.main
Case Else
End Select

End Sub
