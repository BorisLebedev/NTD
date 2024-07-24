Attribute VB_Name = "TableContextMenu"
Sub AddToCellMenu()
Dim TableContextMenu As CommandBar
Dim MySubMenu As CommandBarControl
Dim personal_settings As Boolean
    
Call DeleteFromCellMenu
Set TableContextMenu = Application.CommandBars("Cell")

With TableContextMenu.Controls.Add(Type:=msoControlButton, before:=1)
    .OnAction = "'" & ThisWorkbook.name & "'!" & "TableUpdate.Main"
    .FaceId = 33
    .Caption = "Обновить"
    .Tag = "My_Cell_Control_Tag"
End With

TableContextMenu.Controls(4).BeginGroup = True

End Sub

Sub DeleteFromCellMenu()
Dim TableContextMenu As CommandBar
Dim ctrl As CommandBarControl

Set TableContextMenu = Application.CommandBars("Cell")

For Each ctrl In TableContextMenu.Controls
    If ctrl.Tag = "My_Cell_Control_Tag" Then
        ctrl.Delete
    End If
Next ctrl

On Error GoTo 0
End Sub
