Attribute VB_Name = "WsSubsTimeCollector"
Sub WsSelectionChanged(row As Long, col As Long)

If row > TimeCollector.TOP_INDENT Then
    Select Case col
        Case TimeCollector.COL_LINK
            Call TimeCollector.CalculationLink(row)
        End Select
Else: End If

End Sub


Sub WsSelectionChangedPDF(row As Long, col As Long)

If row > 1 Then
    Select Case col
        Case 4
            Call ValidationOperation(row, col)
        End Select
Else: End If

End Sub


Sub ValidationOperation(row As Long, col As Long)

Cells(row, col).EntireColumn.Validation.Delete

With Cells(row, col).Validation
    .Delete
    .Add Type:=xlValidateList, Operator:=xlBetween, Formula1:="=OPERATIONS"
    .IgnoreBlank = True
    .InCellDropdown = True
    .ShowInput = True
    .ShowError = False
End With

End Sub

