Attribute VB_Name = "SQL"
'Select
Function SqlSelect(sql_str As String, wb As String, wb_path As String, Optional header As String = "no")
Dim CON As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim sql_data As Variant
Dim wb_read_only As Workbook

CON.Provider = "Microsoft.ACE.OLEDB.12.0"
CON.ConnectionString = "data source=" & wb_path & wb & "; extended properties=""Excel 12.0 xml;HDR=" & header & """"
CON.Open
RS.Open sql_str, CON
If Not (RS.BOF And RS.EOF) Then
    sql_data = RS.GetRows
Else: End If
CON.Close

On Error Resume Next
Set wb_read_only = Application.Workbooks(wb)
If Not wb_read_only Is Nothing Then
    wb_read_only.Close
Else: End If

SqlSelect = sql_data
End Function

'UPDATE
Sub SqlUpdate(sql_str As String, wb As String, wb_path As String)
Dim CON As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim wb_read_only As Workbook
On Error GoTo errorline
CON.Provider = "Microsoft.ACE.OLEDB.12.0"
CON.ConnectionString = "data source=" & wb_path & wb & "; extended properties=""Excel 12.0 xml;HDR=no"""
CON.Open
CON.Execute (sql_str)
CON.Close


On Error Resume Next
Set wb_read_only = Application.Workbooks(wb)
If Not wb_read_only Is Nothing Then
    wb_read_only.Close SaveChanges:=False
    MsgBox ("Операция не проведена. Книга открыта другим пользователем.")
Else: End If
Exit Sub
errorline:
    MsgBox ("Операция не проведена")
End Sub


'UPDATE FOR KTTP
Sub SqlUpdateExtended(sql_str_array As Variant, wb As String, wb_path As String)
Dim CON As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim wb_read_only As Workbook
'On Error GoTo errorline
CON.Provider = "Microsoft.ACE.OLEDB.12.0"
CON.ConnectionString = "data source=" & wb_path & wb & "; extended properties=""Excel 12.0 xml;HDR=no"""
CON.Open
For Each sql_str In sql_str_array:
    CON.Execute (sql_str)
Next sql_str
CON.Close


On Error Resume Next
Set wb_read_only = Application.Workbooks(wb)
If Not wb_read_only Is Nothing Then
    wb_read_only.Close SaveChanges:=False
    MsgBox ("Операция не проведена. Книга открыта другим пользователем.")
Else: End If
Exit Sub
errorline:
    MsgBox ("Операция не проведена")
End Sub
