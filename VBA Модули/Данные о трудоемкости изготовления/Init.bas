Attribute VB_Name = "Init"
Public OPERATIONS As Variant
Public OPERATIONS_CORRECTION As Variant
Public OPERATIONS_TYPE_ORDER As Variant


Sub InitConst()
Dim wb As String
Dim wb_path As String
'Dim wb_path As String
Dim sql_str As String
Dim rname As name


wb_path = Paths.OperationsPath
wb = Paths.OperationsName
Call GetOperations(wb, wb_path)
Call GetOperationCorrection(wb, wb_path)
Call GetOperationTypeOrder(wb, wb_path)

End Sub


Private Sub GetOperations(wb As String, wb_path As String)
Dim sql_str As String

sql_str = "SELECT * FROM [Операции$]"

On Error GoTo EXITSUB
OPERATIONS = SQL.SqlSelect(sql_str, wb, wb_path, "yes")

For col = LBound(OPERATIONS) To UBound(OPERATIONS)
    For row = LBound(OPERATIONS, 2) To UBound(OPERATIONS, 2)
        ThisWorkbook.Worksheets("Операции").Cells(row + 1, col + 1) = OPERATIONS(col, row)
    Next row
Next col


OPERATIONS = Calculation.GetOperationArray()

For Each rname In Application.ThisWorkbook.Names
    If rname.name = "OPERATIONS" Then rname.Delete
Next

ThisWorkbook.Names.Add name:="OPERATIONS", RefersTo:="=Операции!$A$1:$A$" & (row)
EXITSUB:
End Sub


Private Sub GetOperationCorrection(wb As String, wb_path As String)
Dim sql_str As String

sql_str = "SELECT * FROM [Исправление$]"
On Error GoTo EXITSUB
OPERATIONS_CORRECTION = SQL.SqlSelect(sql_str, wb, wb_path, "yes")

For col = LBound(OPERATIONS_CORRECTION) To UBound(OPERATIONS_CORRECTION)
    For row = LBound(OPERATIONS_CORRECTION, 2) To UBound(OPERATIONS_CORRECTION, 2)
        ThisWorkbook.Worksheets("Исправления").Cells(row + 1, col + 1) = OPERATIONS_CORRECTION(col, row)
    Next row
Next col

OPERATIONS_CORRECTION = Calculation.GetOperationCorrectionArray()

For Each rname In Application.ThisWorkbook.Names
    If rname.name = "OPERATIONS_CORRECTION" Then rname.Delete
Next

ThisWorkbook.Names.Add name:="OPERATIONS_CORRECTION", RefersTo:="=Исправления!$A$1:$A$" & (row)
EXITSUB:
End Sub


Private Sub GetOperationTypeOrder(wb As String, wb_path As String)
Dim sql_str As String

sql_str = "SELECT * FROM [Порядок видов работ$]"
On Error GoTo EXITSUB
OPERATIONS_TYPE_ORDER = SQL.SqlSelect(sql_str, wb, wb_path, "yes")

For col = LBound(OPERATIONS_TYPE_ORDER) To UBound(OPERATIONS_TYPE_ORDER)
    For row = LBound(OPERATIONS_TYPE_ORDER, 2) To UBound(OPERATIONS_TYPE_ORDER, 2)
        ThisWorkbook.Worksheets("Порядок видов работ").Cells(row + 1, col + 1) = OPERATIONS_TYPE_ORDER(col, row)
    Next row
Next col
ThisWorkbook.Worksheets("Порядок видов работ").Cells(row + 1, 1) = Calculation.OPERATION_ERROR_MSG

OPERATIONS_TYPE_ORDER = Calculation.GetOperationTypeOrderArray()

For Each rname In Application.ThisWorkbook.Names
    If rname.name = "OPERATIONS_TYPE_ORDER" Then rname.Delete
Next

ThisWorkbook.Names.Add name:="OPERATIONS_TYPE_ORDER", RefersTo:="=Порядок видов работ!$A$1:$A$" & (row)
EXITSUB:
End Sub


