Attribute VB_Name = "Module2"
Sub ImportVBAModules()
    Dim vbComp As Object
    Dim importPath As String
    Dim fileName As String
    
    importPath = ThisWorkbook.Path & "\src\"
    
    fileName = Dir(importPath & "*.bas")
    Do While fileName <> ""
        ThisWorkbook.VBProject.VBComponents.Import importPath & fileName
        fileName = Dir
    Loop
    
    MsgBox "Imported VBA Codes! updated", vbInformation
End Sub
