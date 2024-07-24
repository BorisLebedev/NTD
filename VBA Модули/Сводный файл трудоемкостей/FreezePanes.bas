Attribute VB_Name = "FreezePanes"
Private Sub test()
'Range(Cells(1, 1), Cells(1, 318)).EntireColumn.Hidden = False
'Range(Cells(1, 1), Cells(4, 1)).EntireRow.Select
ActiveWindow.SplitRow = 3
ActiveWindow.FreezePanes = True
'Range(Cells(1, 1), Cells(1, 318)).EntireColumn.Hidden = False
'Range(Cells(5, 1), Cells(10000, 375)).AutoFilter

ActiveWindow.SplitColumn = 3
ActiveWindow.FreezePanes = True
End Sub

