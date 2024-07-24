Attribute VB_Name = "ContextMenu"
Sub AddToCellMenu()
Dim ContextMenu As CommandBar
Dim MySubMenu As CommandBarControl
Dim personal_settings As Boolean
    
Call DeleteFromCellMenu
Set ContextMenu = Application.CommandBars("Cell")

With ContextMenu.Controls.Add(Type:=msoControlButton, before:=1)
    .OnAction = "'" & ThisWorkbook.name & "'!" & "DoAll.Main"
    .FaceId = 17
    .Caption = "Полный расчет"
    .Tag = "My_Cell_Control_Tag"
End With

With ContextMenu.Controls.Add(Type:=msoControlButton, before:=2)
    .OnAction = "'" & ThisWorkbook.name & "'!" & "Based.Main"
    .FaceId = 46
    .Caption = "Поиск в Базе"
    .Tag = "My_Cell_Control_Tag"
End With

With ContextMenu.Controls.Add(Type:=msoControlButton, before:=3)
    .OnAction = "'" & ThisWorkbook.name & "'!" & "Levels.LevelsByIndex"
    .FaceId = 11
    .Caption = "Уровни из индексов"
    .Tag = "My_Cell_Control_Tag"
End With

With ContextMenu.Controls.Add(Type:=msoControlButton, before:=4)
    .OnAction = "'" & ThisWorkbook.name & "'!" & "Calculation.Main"
    .FaceId = 33
    .Caption = "Расшифровка"
    .Tag = "My_Cell_Control_Tag"
End With

With ContextMenu.Controls.Add(Type:=msoControlButton, before:=5)
    .OnAction = "'" & ThisWorkbook.name & "'!" & "Consolidation.Main"
    .FaceId = 90
    .Caption = "Консолидация"
    .Tag = "My_Cell_Control_Tag"
End With

With ContextMenu.Controls.Add(Type:=msoControlButton, before:=6)
    .OnAction = "'" & ThisWorkbook.name & "'!" & "Products.Main"
    .FaceId = 984
    .Caption = "Проверка изделий"
    .Tag = "My_Cell_Control_Tag"
End With

With ContextMenu.Controls.Add(Type:=msoControlButton, before:=7)
    .OnAction = "'" & ThisWorkbook.name & "'!" & "Init.InitConst"
    .FaceId = 37
    .Caption = "Обновить данные операций"
    .Tag = "My_Cell_Control_Tag"
End With

ContextMenu.Controls(4).BeginGroup = True

End Sub

Sub DeleteFromCellMenu()
Dim ContextMenu As CommandBar
Dim ctrl As CommandBarControl

Set ContextMenu = Application.CommandBars("Cell")

For Each ctrl In ContextMenu.Controls
    If ctrl.Tag = "My_Cell_Control_Tag" Then
        ctrl.Delete
    End If
Next ctrl

On Error GoTo 0
End Sub
