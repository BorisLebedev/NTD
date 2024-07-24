Attribute VB_Name = "Products"
Const col_level As Integer = 1
Const col_hierarchy As Integer = col_level + 1
Const col_name As Integer = col_hierarchy + 1
Const col_deno As Integer = col_name + 1
Const col_num As Integer = col_deno + 1
Const col_weight As Integer = col_num + 1
Const col_norm As Integer = col_weight + 1
Const col_base As Integer = col_norm + 1
Const l_col As Integer = col_base
Const top_indent As Integer = 1
Const main_ws_name As String = "Расшифровка"


Sub Main()
Const index_col As Integer = Calculation.col_name
Const product_ws_name As String = "Изделия"
Dim data As Variant
Dim ws As Worksheet
Dim l_row As Long
Dim rng As Range
Dim rng_prod As Range

'Call Screen.Events(False)

Set ws = ThisWorkbook.Worksheets(main_ws_name)
With ws
    l_row = DocumentAttribute.LastRow(ws, Calculation.col_name) + 1
    If l_row > Calculation.top_indent Then
        Set rng = .Range(.Cells(Calculation.top_indent + 1, 1), .Cells(l_row, Calculation.l_col))
    Else
        Set rng = .Range(.Cells(Calculation.top_indent + 1, 1), .Cells(Calculation.top_indent + 1, Calculation.l_col))
    End If
End With

data = rng.value
data = Consolidation.CalculateWeights(data)
data_prod = GetProducts(data)

DocumentAttribute.AddNewWS (product_ws_name)
Set ws_product = ThisWorkbook.Worksheets(product_ws_name)

With ws_product
    .Cells.ClearContents
    .Cells.Interior.ColorIndex = 2
    .Cells.Borders.LineStyle = None
    
    .Cells(1, col_hierarchy).EntireColumn.NumberFormat = "@"
    .Cells(1, col_num).EntireColumn.NumberFormat = "0"
    .Cells(1, col_weight).EntireColumn.NumberFormat = "0"
    
    .Cells(1, col_level).EntireColumn.ColumnWidth = 10
    .Cells(1, col_hierarchy).EntireColumn.ColumnWidth = 10
    .Cells(1, col_name).EntireColumn.ColumnWidth = 80
    .Cells(1, col_deno).EntireColumn.ColumnWidth = 20
    .Cells(1, col_norm).EntireColumn.ColumnWidth = 10
    .Cells(1, col_num).EntireColumn.ColumnWidth = 10
    .Cells(1, col_weight).EntireColumn.ColumnWidth = 10
    
    .Cells(1, col_level) = "Уровень"
    .Cells(1, col_hierarchy) = "Индекс"
    .Cells(1, col_name) = "Наименование"
    .Cells(1, col_deno) = "Децимальный номер"
    .Cells(1, col_norm) = "Тр-ть"
    .Cells(1, col_num) = "Кол-во"
    .Cells(1, col_weight) = "Вес"
    .Cells(1, col_base) = "База"
    
    Set rng_prod = .Range(.Cells(top_indent + 1, col_level), .Cells(UBound(data_prod), l_col))
    rng_prod = data_prod
    rng_prod.Borders.LineStyle = XlLineStyle.xlContinuous
    rng_prod.Sort Key1:=ws_product.Cells(top_indent + 1, col_deno)
    
    
    Call SameProductsSameColor(rng_prod)

End With




'Call Screen.Events(True)
End Sub


Private Sub SameProductsSameColor(rng As Range)
Const color1 As Integer = 19
Const color2 As Integer = 2
Const color3 As Integer = 3
Dim color As Integer
Dim deno As String
Dim deno_next As String

color = color1
deno = rng.Rows(1).Cells(1, col_deno).value
rng.Rows(1).Interior.ColorIndex = color

For row = 2 To rng.Rows.Count
    deno_next = rng.Rows(row).Cells(1, col_deno).value
    If deno <> "" Then
        If deno <> deno_next Then
            
            If color = color1 Then
                color = color2
            Else
                color = color1
            End If
            rng.Rows(row).Interior.ColorIndex = color
        Else
            rng.Rows(row).Interior.ColorIndex = color
            If CDec(rng.Rows(row).Cells(1, col_norm).value) <> CDec(rng.Rows(row - 1).Cells(1, col_norm).value) Then
                rng.Rows(row).Cells(1, col_norm).Interior.ColorIndex = color3
                rng.Rows(row - 1).Cells(1, col_norm).Interior.ColorIndex = color3
            Else: End If
        End If
    Else: End If
    deno = deno_next
Next row

End Sub


Private Function GetProducts(data As Variant)

Dim row As Long
Dim row_new As Long
Dim row_sub As Integer
Dim index_hierarchy As String
Dim num As Integer
Dim norm As Double
Dim level As Integer
Dim data_new As Variant

ReDim data_new(1 To l_col, 1 To 1)
l_row = UBound(data)
For row = LBound(data) To UBound(data)
    index_hierarchy = data(row, Calculation.col_hierarchy)
    num = data(row, Calculation.col_num)
    If index_hierarchy <> "" And num <> 0 Then
        
        row_new = UBound(data_new, 2)
        ReDim Preserve data_new(LBound(data_new) To UBound(data_new), _
                                LBound(data_new, 2) To UBound(data_new, 2) + 1)
        
        data_new(col_level, row_new) = data(row, Calculation.col_level)
        data_new(col_hierarchy, row_new) = data(row, Calculation.col_hierarchy)
        data_new(col_name, row_new) = data(row, Calculation.col_name)
        data_new(col_deno, row_new) = data(row, Calculation.col_deno)
        data_new(col_norm, row_new) = data(row, Calculation.col_norm_calc) / data(row, Calculation.col_num)
        data_new(col_num, row_new) = data(row, Calculation.col_num)
        data_new(col_weight, row_new) = data(row, Calculation.col_deno_td)
        data_new(col_base, row_new) = data(row, Calculation.col_base)
    Else: End If
Next
data_new = DocumentAttribute.Transpose2dArray(data_new)


GetProducts = data_new
End Function

