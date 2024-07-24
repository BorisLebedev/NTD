Attribute VB_Name = "Update"
Private Sub update()

Dim WBNew As Workbook
Dim WBOld As Workbook
Dim path As String
Dim path_modules As String
Dim path_upd As String
Dim oFSO As Object
Dim oFolder As Object
Dim oFile As Object

Call Screen.Events(False)

path = Application.ThisWorkbook.path & "\"
path_modules = path & "VBA Модули\"
path_upd = path & "Обновляемые расшифровки\"

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFolder = oFSO.GetFolder(path_upd)

Set WBNew = ThisWorkbook()
Call ExportModules(WBNew, path_modules)

For Each oFile In oFolder.Files
    If Right(oFile.name, Len(WBNew.name)) <> WBNew.name Then
        Set WBOld = Workbooks.Open(path_upd & oFile.name)
        Application.StatusBar = WBOld.name
        Call CopyModule(WBOld, WBNew, path_modules)
        WBOld.Save
        WBOld.Close
    Else: End If
Next

Call Screen.Events(True)
Application.StatusBar = "Обновление завершено"

End Sub


Private Sub CopyModule(WBOld As Workbook, WBNew As Workbook, path_modules As String)

Call DeleteModules(WBOld)
Call AddModules(WBOld, WBNew, path_modules)

End Sub


Private Sub DeleteModules(WBOld As Workbook)

For Each component In WBOld.VBProject.VBComponents
    If component.Type = 1 Then
        WBOld.VBProject.VBComponents.Remove component
    Else: End If
Next

End Sub


Private Sub ExportModules(wb As Workbook, path As String)

For Each component In wb.VBProject.VBComponents
    If component.Type = 1 Then
        str_export = path & component.name & ".bas"
        component.Export str_export
    Else: End If
Next

End Sub

Private Sub AddModules(WBOld As Workbook, WBNew As Workbook, path_modules As String)
Dim oFSO As Object
Dim oFolder As Object
Dim oFile As Object

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFolder = oFSO.GetFolder(path_modules)

For Each oFile In oFolder.Files
    WBOld.VBProject.VBComponents.Import oFile
Next

End Sub
