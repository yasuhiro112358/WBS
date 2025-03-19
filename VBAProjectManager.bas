Attribute VB_Name = "VBAProjectManager"

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

Sub ImportVBAModules()
    Dim vbComp As Object
    Dim importPath As String
    Dim fileName As String
    
    importPath = ThisWorkbook.Path & "\src\"
    
    fileName = Dir(importPath & "*.bas")
    Do While fileName <> ""
        moduleName = Left(fileName, InStrRev(fileName, ".") - 1)

        On Error Resume Next
        ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(moduleName)
        On Error GoTo 0

        ThisWorkbook.VBProject.VBComponents.Import importPath & fileName
        fileName = Dir
    Loop
        
    MsgBox "Imported VBA Codes!", vbInformation
End Sub

