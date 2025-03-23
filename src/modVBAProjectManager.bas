Attribute VB_Name = "modVBAProjectManager"
Option Explicit

Sub ExportVBAModules()
    Dim vbComp As Object
    Dim exportPath As String
    Dim userResponse As VbMsgBoxResult

    exportPath = ThisWorkbook.Path & Application.PathSeparator & "src" & Application.PathSeparator
  

    ' Confirmation message
    userResponse = MsgBox( _
        "Do you want to export VBA modules?" & vbCrLf & vbCrLf & _
        "Destination Folder:" & vbCrLf & _
        exportPath & vbCrLf & vbCrLf & _
        "Click [Yes] to proceed or [No] to cancel.", _
        vbYesNo + vbQuestion, "Confirm VBA Export")
    If userResponse = vbNo Then
        Exit Sub
    End If
  
    ' Create the folder if it does not exist
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If
    
    ' Export each component
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        ExportModule vbComp, exportPath
    Next vbComp
    
    MsgBox "VBA code has been exported!", vbInformation
End Sub

' Export a module with the appropriate file extension
Private Sub ExportModule(ByVal vbComp As Object, ByVal exportPath As String)
    Dim fileExtension As String
    Select Case vbComp.Type
        Case 1 ' vbext_ct_StdModule
            fileExtension = ".bas"
        Case 2 ' vbext_ct_ClassModule
            fileExtension = ".cls"
        Case 3 ' vbext_ct_MSForm
            fileExtension = ".frm"
        Case 100 ' vbext_ct_Document
            fileExtension = ".cls"
        Case Else
            Exit Sub
    End Select
    
    vbComp.Export exportPath & vbComp.Name & fileExtension
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







