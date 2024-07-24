Attribute VB_Name = "ContextMenu"
Sub AddToCellMenu()
Dim ContextMenu As CommandBar
Dim MySubMenu As CommandBarControl
Dim personal_settings As Boolean
    
Call DeleteFromCellMenu
Set ContextMenu = Application.CommandBars("Cell")
'Set SubContextUpd = ContextMenu.Controls.Add(Type:=msoControlPopup, before:=1)
'Set SubContextMenuR = ContextMenu.Controls.Add(Type:=msoControlPopup, before:=3)
'Set SubContextMenuRCable = SubContextMenuR.Controls.Add(Type:=msoControlPopup, before:=1)
'Set SubContextMenuRAssambly = SubContextMenuR.Controls.Add(Type:=msoControlPopup, before:=2)
'Set SubContextMenuRNode = SubContextMenuR.Controls.Add(Type:=msoControlPopup, before:=2)

'Set SubContextMenuZ = ContextMenu.Controls.Add(Type:=msoControlPopup, before:=4)
'Set SubContextMenuPKI = SubContextMenuZ.Controls.Add(Type:=msoControlPopup, before:=1)
'Set SubContextMenuSTC = SubContextMenuZ.Controls.Add(Type:=msoControlPopup, before:=2)



With ContextMenu.Controls.Add(Type:=msoControlButton, before:=1)
    .OnAction = "'" & ThisWorkbook.name & "'!" & "Calculation.Main"
    .FaceId = 17
    .Caption = "Полный расчет"
    .Tag = "My_Cell_Control_Tag"
End With

With ContextMenu.Controls.Add(Type:=msoControlButton, before:=2)
    .OnAction = "'" & ThisWorkbook.name & "'!" & "Levels.LevelsByIndex"
    .FaceId = 11
    .Caption = "Уровни из индексов"
    .Tag = "My_Cell_Control_Tag"
End With

With ContextMenu.Controls.Add(Type:=msoControlButton, before:=3)
    .OnAction = "'" & ThisWorkbook.name & "'!" & "ExportData.Main"
    .Caption = "В форму НТД"
    .Tag = "My_Cell_Control_Tag"
End With


'With SubContextMenuR
'
'    .Caption = "Ремонт"
'    .Tag = "New_Item_Context_Menu"
'
'    With .Controls.Add(Type:=msoControlButton)
'        '.OnAction = "'" & ThisWorkbook.Name & "'!" & "LowerMacro"
'        .OnAction = "'" & ThisWorkbook.name & "'!" & "ContextMenu.GetType_1"
'        '.FaceId = 91
'        .Caption = "Комплекс"
'
'    End With
'
'    With .Controls.Add(Type:=msoControlButton)
'        '.OnAction = "'" & ThisWorkbook.Name & "'!" & "ProperMacro"
'        .OnAction = "'" & ThisWorkbook.name & "'!" & "ContextMenu.GetType_2"
'        '.FaceId = 95
'        .Caption = "Комплект"
'    End With
'
'End With
'
'
'With SubContextMenuRCable
'
'    .Caption = "Кабель"
'    .Tag = "New_Item_Context_Menu"
'
'    With .Controls.Add(Type:=msoControlButton)
'        '.OnAction = "'" & ThisWorkbook.Name & "'!" & "UpperMacro"
'        .OnAction = "'" & ThisWorkbook.name & "'!" & "ContextMenu.GetType_3"
'        '.FaceId = 100
'        .Caption = "Кабель"
'    End With
'
'    With .Controls.Add(Type:=msoControlButton)
'        '.OnAction = "'" & ThisWorkbook.Name & "'!" & "LowerMacro"
'        .OnAction = "'" & ThisWorkbook.name & "'!" & "ContextMenu.GetType_4"
'        '.FaceId = 91
'        .Caption = "Кабельная сборка"
'    End With
'
'End With
'
'With SubContextMenuRAssambly
'
'    .Caption = "Сборка"
'    .Tag = "New_Item_Context_Menu"
'
'    With .Controls.Add(Type:=msoControlButton)
'        '.OnAction = "'" & ThisWorkbook.Name & "'!" & "UpperMacro"
'        .OnAction = "'" & ThisWorkbook.name & "'!" & "ContextMenu.GetType_5"
'        '.FaceId = 100
'        .Caption = "С настройкой"
'    End With
'
'    With .Controls.Add(Type:=msoControlButton)
'        '.OnAction = "'" & ThisWorkbook.Name & "'!" & "LowerMacro"
'        .OnAction = "'" & ThisWorkbook.name & "'!" & "ContextMenu.GetType_6"
'        '.FaceId = 91
'        .Caption = "Без настройки"
'    End With
'
'End With
'
'With SubContextMenuRNode
'
'    .Caption = "Узел"
'    .Tag = "New_Item_Context_Menu"
'
'    With .Controls.Add(Type:=msoControlButton)
'        '.OnAction = "'" & ThisWorkbook.Name & "'!" & "LowerMacro"
'        .OnAction = "'" & ThisWorkbook.name & "'!" & "ContextMenu.GetType_7"
'        '.FaceId = 91
'        .Caption = "Механический"
'    End With
'
'    With .Controls.Add(Type:=msoControlButton)
'        '.OnAction = "'" & ThisWorkbook.Name & "'!" & "LowerMacro"
'        .OnAction = "'" & ThisWorkbook.name & "'!" & "ContextMenu.GetType_8"
'        '.FaceId = 91
'        .Caption = "Электрический"
'    End With
'End With
'
'With SubContextMenuZ
'
'    .Caption = "Замена"
'    .Tag = "New_Item_Context_Menu"
'
'End With
'
'With SubContextMenuPKI
'
'    .Caption = "ПКИ"
'    .Tag = "New_Item_Context_Menu"
'
'    With .Controls.Add(Type:=msoControlButton)
'        '.OnAction = "'" & ThisWorkbook.Name & "'!" & "UpperMacro"
'        .OnAction = "'" & ThisWorkbook.name & "'!" & "ContextMenu.GetType_9"
'        '.FaceId = 100
'        .Caption = "ПКИ"
'    End With
'
'    With .Controls.Add(Type:=msoControlButton)
'        '.OnAction = "'" & ThisWorkbook.Name & "'!" & "LowerMacro"
'        .OnAction = "'" & ThisWorkbook.name & "'!" & "ContextMenu.GetType_10"
'        '.FaceId = 91
'        .Caption = "Деталь (ПКИ)"
'    End With
'
'End With
'
'With SubContextMenuSTC
'
'    .Caption = "СТЦ"
'    .Tag = "New_Item_Context_Menu"
'
'    With .Controls.Add(Type:=msoControlButton)
'        '.OnAction = "'" & ThisWorkbook.Name & "'!" & "UpperMacro"
'        .OnAction = "'" & ThisWorkbook.name & "'!" & "ContextMenu.GetType_11"
'        '.FaceId = 100
'        .Caption = "Деталь"
'    End With
'
'    With .Controls.Add(Type:=msoControlButton)
'        '.OnAction = "'" & ThisWorkbook.Name & "'!" & "LowerMacro"
'        .OnAction = "'" & ThisWorkbook.name & "'!" & "ContextMenu.GetType_12"
'        '.FaceId = 91
'        .Caption = "Плата или ячейка"
'    End With
'
'    With .Controls.Add(Type:=msoControlButton)
'        '.OnAction = "'" & ThisWorkbook.Name & "'!" & "LowerMacro"
'        .OnAction = "'" & ThisWorkbook.name & "'!" & "ContextMenu.GetType_13"
'        '.FaceId = 91
'        .Caption = "СПО"
'    End With
'
'End With

ContextMenu.Controls(4).BeginGroup = True

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
    On Error GoTo PASS
        ctrl.Delete
PASS:
    End If
Next ctrl

On Error GoTo 0
End Sub

Private Sub GetType_1()
Dim product_type As String
product_type = ThisWorkbook.Worksheets("Типы").Cells(9, 4).Value2
SetType (product_type)
End Sub

Private Sub GetType_2()
product_type = ThisWorkbook.Worksheets("Типы").Cells(10, 4).Value2
SetType (product_type)
End Sub

Private Sub GetType_3()
product_type = ThisWorkbook.Worksheets("Типы").Cells(7, 4).Value2
SetType (product_type)
End Sub

Private Sub GetType_4()
product_type = ThisWorkbook.Worksheets("Типы").Cells(8, 4).Value2
SetType (product_type)
End Sub

Private Sub GetType_5()
product_type = ThisWorkbook.Worksheets("Типы").Cells(12, 4).Value2
SetType (product_type)
End Sub

Private Sub GetType_6()
product_type = ThisWorkbook.Worksheets("Типы").Cells(11, 4).Value2
SetType (product_type)
End Sub

Private Sub GetType_7()
product_type = ThisWorkbook.Worksheets("Типы").Cells(13, 4).Value2
SetType (product_type)
End Sub

Private Sub GetType_8()
product_type = ThisWorkbook.Worksheets("Типы").Cells(14, 4).Value2
SetType (product_type)
End Sub

Private Sub GetType_9()
product_type = ThisWorkbook.Worksheets("Типы").Cells(3, 4).Value2
SetType (product_type)
End Sub

Private Sub GetType_10()
product_type = ThisWorkbook.Worksheets("Типы").Cells(2, 4).Value2
SetType (product_type)
End Sub

Private Sub GetType_11()
product_type = ThisWorkbook.Worksheets("Типы").Cells(4, 4).Value2
SetType (product_type)
End Sub

Private Sub GetType_12()
product_type = ThisWorkbook.Worksheets("Типы").Cells(5, 4).Value2
SetType (product_type)
End Sub

Private Sub GetType_13()
product_type = ThisWorkbook.Worksheets("Типы").Cells(6, 4).Value2
SetType (product_type)
End Sub

Private Sub SetType(product_type As String)
Dim row_calc As Long
ActiveCell = product_type
row_calc = ActiveCell.row
TypeOfProduct.SetBaseValue (row_calc)
End Sub

