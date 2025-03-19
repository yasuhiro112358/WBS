Attribute VB_Name = "Module2"
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
