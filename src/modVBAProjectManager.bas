Attribute VB_Name = "modVBAProjectManager"
Option Explicit

Sub ExportVBAModules()
    Dim vbComp As Object
    Dim exportPath As String
    
    exportPath = ThisWorkbook.Path & Application.PathSeparator & "src" & Application.PathSeparator
    
    ' Create the folder if it does not exist
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If
    
    ' Export each module (excluding VBAProjectManager)
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Name <> "VBAProjectManager" Then
            Select Case vbComp.Type
                Case 1 ' Standard module
                    vbComp.Export exportPath & vbComp.Name & ".bas"
                Case 2 ' Class module
                    vbComp.Export exportPath & vbComp.Name & ".cls"
                Case 3 ' UserForm
                    vbComp.Export exportPath & vbComp.Name & ".frm"
            End Select
        End If
    Next vbComp
    
    MsgBox "VBA code has been exported! (excluding VBAProjectManager)", vbInformation
End Sub

Sub ImportVBAModules()
    Dim vbComp As Object
    Dim importPath As String
    Dim fileName As String
    Dim moduleName As String
    
    importPath = ThisWorkbook.Path & Application.PathSeparator & "src" & Application.PathSeparator

    If Not IsVBProjectAccessible() Then
        MsgBox GetVBATrustAccessMessage(), vbCritical, "VBA Project Access Error"
        Exit Sub
    End If

    ' Import .bas files
    fileName = Dir(importPath & "*.bas")
    Do While fileName <> ""
        moduleName = Left(fileName, InStrRev(fileName, ".") - 1)
        
        ' Do not delete or import VBAProjectManager
        If moduleName <> "VBAProjectManager" Then
            ' Delete existing module
            On Error Resume Next
            Set vbComp = ThisWorkbook.VBProject.VBComponents(moduleName)
            If Not vbComp Is Nothing Then
                ThisWorkbook.VBProject.VBComponents.Remove vbComp
            End If
            Err.Clear
            On Error GoTo 0

            ' Import
            ThisWorkbook.VBProject.VBComponents.Import importPath & fileName
        End If

        fileName = Dir
    Loop

    ' Import .cls files
    fileName = Dir(importPath & "*.cls")
    Do While fileName <> ""
        moduleName = Left(fileName, InStrRev(fileName, ".") - 1)
        
        ' Delete existing class module
        On Error Resume Next
        Set vbComp = ThisWorkbook.VBProject.VBComponents(moduleName)
        If Not vbComp Is Nothing Then
            ThisWorkbook.VBProject.VBComponents.Remove vbComp
        End If
        Err.Clear
        On Error GoTo 0

        ' Import
        ThisWorkbook.VBProject.VBComponents.Import importPath & fileName
        fileName = Dir
    Loop

    ' Import .frm files (UserForms)
    fileName = Dir(importPath & "*.frm")
    Do While fileName <> ""
        moduleName = Left(fileName, InStrRev(fileName, ".") - 1)
        
        ' Delete existing UserForm
        On Error Resume Next
        Set vbComp = ThisWorkbook.VBProject.VBComponents(moduleName)
        If Not vbComp Is Nothing Then
            ThisWorkbook.VBProject.VBComponents.Remove vbComp
        End If
        Err.Clear
        On Error GoTo 0

        ' Import
        ThisWorkbook.VBProject.VBComponents.Import importPath & fileName
        fileName = Dir
    Loop

    MsgBox "VBA code has been imported!", vbInformation
End Sub

Function IsVBProjectAccessible() As Boolean
    Dim Test As Object

    On Error Resume Next
    Set Test = ThisWorkbook.VBProject
    IsVBProjectAccessible = (Err.Number = 0)
    On Error GoTo 0
End Function

Function GetVBATrustAccessMessage() As String
    GetVBATrustAccessMessage = _
        "This VBA project is not accessible. " & vbCrLf & _
        "To enable access to the VBA project object model, follow these steps:" & vbCrLf & vbCrLf & _
        "1. Open Excel and click [File] > [Options]." & vbCrLf & _
        "2. Go to [Trust Center] and click [Trust Center Settings]." & vbCrLf & _
        "3. Select [Macro Settings] and check " & _
        "'Trust access to the VBA project object model'." & vbCrLf & _
        "4. Restart Excel to apply the changes." & vbCrLf & vbCrLf & _
        "After applying these settings, run this macro again."
End Function





