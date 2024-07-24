Attribute VB_Name = "ContextMenu"
Sub AddToCellMenu()
Dim ContextMenu As CommandBar
Dim MySubMenu As CommandBarControl
Dim personal_settings As Boolean
    
Call DeleteFromCellMenu
Set ContextMenu = Application.CommandBars("Cell")

With ContextMenu.Controls.Add(Type:=msoControlButton, before:=1)
    .OnAction = "'" & ThisWorkbook.name & "'!" & "Main.main"
    .FaceId = 17
    .Caption = "Обновить"
    .Tag = "New_Item_Context_Menu"
End With

With ContextMenu.Controls.Add(Type:=msoControlButton, before:=2)
    .OnAction = "'" & ThisWorkbook.name & "'!" & "Time.time"
    .FaceId = 17
    .Caption = "Данные из расшифровок"
    .Tag = "New_Item_Context_Menu"
End With

End Sub


Sub DeleteFromCellMenu()
Dim ContextMenu As CommandBar
Dim ctrl As CommandBarControl

Set ContextMenu = Application.CommandBars("Cell")

For Each ctrl In ContextMenu.Controls
    If ctrl.Tag = "New_Item_Context_Menu" Then
        ctrl.Delete
    End If
Next ctrl

For Each ctrl In ContextMenu.Controls
    If ctrl.Tag = "My_Cell_Control_Tag" Then
        ctrl.Delete
    End If
Next ctrl

On Error GoTo 0
End Sub

