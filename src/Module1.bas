Attribute VB_Name = "Module1"
Sub ExportVBAModules()
    Dim vbComp As Object
    Dim exportPath As String
    
    exportPath = ThisWorkbook.Path & "\src\"
    
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1, 2, 3 ' Module, Class Module, UserForm
                vbComp.Export exportPath & vbComp.Name & ".bas"
        End Select
    Next vbComp
    
    MsgBox "Exported VBA Codes!", vbInformation
End Sub
