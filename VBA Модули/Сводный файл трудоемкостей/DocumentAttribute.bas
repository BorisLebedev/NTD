Attribute VB_Name = "DocumentAttribute"
'��� ���� ��
Function TdTypeDoc(td_denotation_number As String)
Select Case Right(Left(td_denotation_number, 7), 2)
    Case Is = "25"
        TdTypeDoc = "��"
    Case Is = "10"
        TdTypeDoc = "��"
    Case Is = "50"
        TdTypeDoc = "���"
    Case Is = "55"
        TdTypeDoc = "��(�)��"
    Case Is = "20"
        TdTypeDoc = "��"
    Case Is = "60"
        TdTypeDoc = "��"
    Case Is = "57"
        TdTypeDoc = "���"
    Case Is = "75"
        TdTypeDoc = "���"
    Case Is = "30"
        TdTypeDoc = "��"
    Case Is = "62"
        TdTypeDoc = "��"
    Case Is = "59"
        TdTypeDoc = "���"
    Case Is = "67"
        TdTypeDoc = "���"
    Case Is = "66"
        TdTypeDoc = "���"
    Case Is = "01"
        TdTypeDoc = "��"
    Case Is = "02"
        TdTypeDoc = "��(�)"
    Case Is = "04"
        TdTypeDoc = "��(���)"
    Case Is = "05"
        TdTypeDoc = "��(��)"
    Case Is = "06"
        TdTypeDoc = "��(�)"
    Case Is = "07"
        TdTypeDoc = "��(��)"
    Case Is = "09"
        TdTypeDoc = "��(���)"
    Case Is = "80"
        TdTypeDoc = "���"
    Case Is = "40"
        TdTypeDoc = "���"
    Case Is = "41"
        TdTypeDoc = "��"
    Case Is = "72"
        TdTypeDoc = "���"
    Case Is = "45"
        TdTypeDoc = "���"
    Case Is = "71"
        TdTypeDoc = "��"
    Case Is = "00"
        TdTypeDoc = "���"
    Case Is = "43"
        TdTypeDoc = "��"
    Case Is = "48"
        TdTypeDoc = "���"
    Case Is = "47"
        TdTypeDoc = "���"
    Case Is = "46"
        TdTypeDoc = "���"
    Case Is = "42"
        TdTypeDoc = "��"
    Case Is = "44"
        TdTypeDoc = "���"
    Case Is = "78"
        TdTypeDoc = "��"
    Case Is = "77"
        TdTypeDoc = "���"
    Case Is = "79"
        TdTypeDoc = "���"
    Case Is = "70"
        TdTypeDoc = "��"
    Case Else
End Select
End Function


'��� ���� �� �� �����������
Function TdTypeOrg(td_denotation_number As String)
Select Case Right(Left(td_denotation_number, 8), 1)
     Case Is = "1"
     TdTypeOrg = "�"
     Case Is = "2"
     TdTypeOrg = "�"
     Case Is = "3"
     TdTypeOrg = "�"
     Case Is = "0"
     TdTypeOrg = "�/�"
     Case Else
End Select
End Function


'��� ���� �� �� ������
Function TdTypeMethod(td_denotation_number As String)
Select Case Right(Left(td_denotation_number, 10), 2)
     Case Is = "00"
     TdTypeMethod = "��� ��������"
     Case Is = "88"
     TdTypeMethod = "������"
     Case Is = "01"
     TdTypeMethod = "������ ����������"
     Case Is = "85"
     TdTypeMethod = "�������������"
     Case Is = "80"
     TdTypeMethod = "�����"
     Case Is = "02"
     TdTypeMethod = "����������� ��������"
     Case Is = "21"
     TdTypeMethod = "��������� ���������"
     Case Is = "41"
     TdTypeMethod = "��������� ��������"
     Case Is = "06"
     TdTypeMethod = "���������"
     Case Is = "08"
     TdTypeMethod = "����������� � ������������"
     Case Is = "73"
     TdTypeMethod = "��������� �������� ������������� (������������)"
     Case Is = "71"
     TdTypeMethod = "��������� �������� (�������������� � ���������������� ���������������)"
     Case Is = "75"
     TdTypeMethod = "�����������������, ����������������� � ������������ ���������"
     Case Is = "65"
     TdTypeMethod = "���������� �����������"
     Case Is = "55"
     TdTypeMethod = "����������-���������� ���������"
     Case Is = "60"
     TdTypeMethod = "���������������� �� ���������� ����������, ��������, ������ � ������"
     Case Is = "50"
     TdTypeMethod = "��������������"
     Case Is = "10"
     TdTypeMethod = "����� �������� � �������"
     Case Is = "04"
     TdTypeMethod = "�����������"
     Case Is = "90"
     TdTypeMethod = "������"
     Case Else
End Select
End Function

'����� ������ ���������� ����
Function LastRow(ws As Worksheet, column As Integer)
Dim l_row As Long
l_row = ws.Cells(1048576, column).End(xlUp).row
LastRow = l_row
End Function

'��� �� �������� ���������
Function TdTypeDocCode(doc_name As String)
Dim code As String

Select Case doc_name
Case Is = "��������������� ����������"
    code = "25"
Case Is = "���������� �����"
    code = "10"
Case Is = "����� ���������������� ��������"
    code = "50"
Case Is = "����� �������� (����������) ���������������� ��������"
    code = "55"
Case Is = "����� �������"
    code = "20"
Case Is = "������������ �����"
    code = "60"
Case Is = "����� ������� (���������) ��������"
    code = "57"
Case Is = "�������-������������� �����"
    code = "75"
Case Is = "��������������� �����"
    code = "30"
Case Is = "����� �������"
    code = "62"
Case Is = "����� ��������������� ����������"
    code = "59"
Case Is = "����� ����������� ����������"
    code = "67"
Case Is = "����� ������� ����������"
    code = "66"
Case Is = "�������� ��������������� ������������"
    code = "01"
Case Is = "�������� ���������� �� (��������)"
    code = "02"
Case Is = "�������� ��������� ���������� �� (��������)"
    code = "04"
Case Is = "�������� ��������� ��������������� ������������"
    code = "05"
Case Is = "�������� ����������� ��������������� ������������"
    code = "06"
Case Is = "�������� ���������� �� (��������) ��������������� ����������"
    code = "07"
Case Is = "����������� �������� ���������� �� (��������)"
    code = "09"
Case Is = "��������� ���������� �����������"
    code = "80"
Case Is = "��������� ��������������� ����������"
    code = "40"
Case Is = "��������� ��������������� ���������"
    code = "41"
Case Is = "��������� ��������"
    code = "72"
Case Is = "��������� ������ �������"
    code = "45"
Case Is = "��������� �������������"
    code = "71"
Case Is = "��������� ������������ �����"
    code = "00"
Case Is = "��������� ����������"
    code = "43"
Case Is = "��������� �������� ���� ������� ����������"
    code = "48"
Case Is = "��������� ����������������� ���� ������� ����������"
    code = "47"
Case Is = "��������� ������������"
    code = "46"
Case Is = "��������� ��������"
    code = "42"
Case Is = "��������� ��� � �������� (����������) �� (��������)"
    code = "44"
Case Is = "��������� ����������"
    code = "78"
Case Is = "��������� �������, ������������� �� �������"
    code = "77"
Case Is = "��������� ��������"
    code = "79"
Case Is = "��������������� ���������"
    code = "70"
Case Else
    code = ""
End Select

TdTypeDocCode = code
End Function


'���������� �������� ���������
Function TdTypeDocAbbrev(doc_name As String)
Dim code As String

Select Case doc_name
Case Is = "��������������� ����������"
    code = "��"
Case Is = "���������� �����"
    code = "��"
Case Is = "����� ���������������� ��������"
    code = "���"
Case Is = "����� �������� (����������) ���������������� ��������"
    code = "��(�)��"
Case Is = "����� �������"
    code = "��"
Case Is = "������������ �����"
    code = "��"
Case Is = "����� ������� (���������) ��������"
    code = "���"
Case Is = "�������-������������� �����"
    code = "���"
Case Is = "��������������� �����"
    code = "��"
Case Is = "����� �������"
    code = "��"
Case Is = "����� ��������������� ����������"
    code = "���"
Case Is = "����� ����������� ����������"
    code = "���"
Case Is = "����� ������� ����������"
    code = "���"
Case Is = "�������� ��������������� ������������"
    code = "��"
Case Is = "�������� ���������� �� (��������)"
    code = "��(�)"
Case Is = "�������� ��������� ���������� �� (��������)"
    code = "��(���)"
Case Is = "�������� ��������� ��������������� ������������"
    code = "��(��)"
Case Is = "�������� ����������� ��������������� ������������"
    code = "��(�)"
Case Is = "�������� ���������� �� (��������) ��������������� ����������"
    code = "��(��)"
Case Is = "����������� �������� ���������� �� (��������)"
    code = "��(���)"
Case Is = "��������� ���������� �����������"
    code = "���"
Case Is = "��������� ��������������� ����������"
    code = "���"
Case Is = "��������� ��������������� ���������"
    code = "��"
Case Is = "��������� ��������"
    code = "���"
Case Is = "��������� ������ �������"
    code = "���"
Case Is = "��������� �������������"
    code = "��"
Case Is = "��������� ������������ �����"
    code = "���"
Case Is = "��������� ����������"
    code = "��"
Case Is = "��������� �������� ���� ������� ����������"
    code = "���"
Case Is = "��������� ����������������� ���� ������� ����������"
    code = "���"
Case Is = "��������� ������������"
    code = "���"
Case Is = "��������� ��������"
    code = "��"
Case Is = "��������� ��� � �������� (����������) �� (��������)"
    code = "���"
Case Is = "��������� ����������"
    code = "��"
Case Is = "��������� �������, ������������� �� �������"
    code = "���"
Case Is = "��������� ��������"
    code = "���"
Case Is = "��������������� ���������"
    code = "��"
Case Else
    code = ""
End Select

TdTypeDocAbbrev = code
End Function


'��� �� ���� �����������
Function TdTypeOrgCode(doc_name As String)
Dim code As String

Select Case UserFormBase.VidTDOrg.value
Case Is = "��������� ������� (��������)"
    code = "1"
Case Is = "������� ������� (��������)"
    code = "2"
Case Is = "������� ������� (��������)"
    code = "3"
Case Is = "��� ��������"
    code = "0"
Case Else
    code = ""
End Select

TdTypeOrgCode = code
End Function


'���������� �� ���� �����������
Function TdTypeOrgAbbrev(doc_name As String)
Dim code As String

Select Case UserFormBase.VidTDOrg.value
Case Is = "��������� ������� (��������)"
    code = "�"
Case Is = "������� ������� (��������)"
    code = "�"
Case Is = "������� ������� (��������)"
    code = "�"
Case Is = "��� ��������"
    code = "�/�"
Case Else
    code = ""
End Select

TdTypeOrgAbbrev = code
End Function

Function TdTypeMethodCode(doc_name As String)
Dim code As String

Select Case doc_name
Case Is = "��� ��������"
    code = "00"
Case Is = "������"
    code = "88"
Case Is = "������ ����������"
    code = "01"
Case Is = "�������������"
    code = "85"
Case Is = "�����"
    code = "80"
Case Is = "����������� ��������"
    code = "02"
Case Is = "��������� ���������"
    code = "21"
Case Is = "��������� ��������"
    code = "41"
Case Is = "���������"
    code = "06"
Case Is = "����������� � ������������"
    code = "08"
Case Is = "��������� �������� ������������� (������������)"
    code = "73"
Case Is = "��������� �������� (�������������� � ���������������� ���������������)"
    code = "71"
Case Is = "�����������������, ����������������� � ������������ ���������"
    code = "75"
Case Is = "���������� �����������"
    code = "65"
Case Is = "����������-���������� ���������"
    code = "55"
Case Is = "���������������� �� ���������� ����������, ��������, ������ � ������"
    code = "60"
Case Is = "��������������"
    code = "50"
Case Is = "����� �������� � �������"
    code = "10"
Case Is = "�����������"
    code = "04"
Case Is = "������"
    code = "90"
Case Else
    code = ""
End Select

TdTypeMethodCode = code
End Function


Function InArray(arr As Variant, value As String)
InArray = False
For Each arr_val In arr
    If value = arr_val Then
        InArray = True
        Exit For
    Else: End If
Next arr_val
End Function

Sub AddNewWS(name As String)
Dim ws As Worksheet
Dim exist As Boolean

exist = False
exist = WorksheetExists(name)

If exist Then
    Set ws = ThisWorkbook.Worksheets(name)
Else
    With ThisWorkbook
        Set ws = .sheets.Add(After:=.sheets(.sheets.Count))
        ws.name = name
    End With
End If

End Sub


Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
Dim sht As Worksheet

If wb Is Nothing Then Set wb = ThisWorkbook
On Error Resume Next
Set sht = wb.sheets(shtName)
On Error GoTo 0
WorksheetExists = Not sht Is Nothing
End Function


Function Reverse2dArray(data As Variant)
Dim data_new As Variant
Dim row_new As Long
data_new = data
row_new = LBound(data)
For row = UBound(data) To 1 Step -1
    For col = LBound(data, 2) To UBound(data, 2)
        data_new(row_new, col) = data(row, col)
    Next col
    row_new = row_new + 1
Next row
Reverse2dArray = data_new
End Function


Function Transpose2dArray(arr As Variant)
ReDim new_arr(LBound(arr, 2) To UBound(arr, 2), LBound(arr) To UBound(arr))
For row = LBound(new_arr) To UBound(new_arr)
    For col = LBound(new_arr, 2) To UBound(new_arr, 2)
        new_arr(row, col) = arr(col, row)
    Next col
Next row
Transpose2dArray = new_arr
End Function


Function FileCounter(sPath As String)
Dim doc_counter As Integer

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFolder = oFSO.GetFolder(sPath)
 
doc_counter = 1
'���� ����������� ���������� ������ � �����
For Each oFile In oFolder.Files
    Application.StatusBar = "���������� ���������� = " & doc_counter
    wb_name = oFile.name
    doc_counter = doc_counter + 1
Next
        
FileCounter = doc_counter
End Function
