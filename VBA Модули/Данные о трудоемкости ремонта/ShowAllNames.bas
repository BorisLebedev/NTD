Attribute VB_Name = "ShowAllNames"
Sub ShowAllNames()
    For Each n In ThisWorkbook.Names
        n.Visible = True
    Next
End Sub
