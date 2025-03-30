Attribute VB_Name = "modVBAProjectManager"
Option Explicit

Sub ExportVBAModules()
    Dim Sep As String
    Sep = Application.PathSeparator
    
    Dim ExportDir As String
    ExportDir = ThisWorkbook.Path & Sep & "exported"
    
    Dim UserConfirmed As Boolean
    UserConfirmed = ConfirmExport(ExportDir)
    If Not UserConfirmed Then
        Exit Sub
    End If
    
    If Dir(ExportDir, vbDirectory) = "" Then
        MkDir ExportDir
    End If
    
    Dim objComponent As Object
    For Each objComponent In ThisWorkbook.VBProject.VBComponents
        ExportModule objComponent, ExportDir & Sep 'Need separator
    Next objComponent
    
    MsgBox "VBA code has been exported!", vbInformation
    Call LogMessage("VBA code has been exported.")
End Sub

Private Function ConfirmExport(p_Destination As String) As Boolean
    Dim Response As VbMsgBoxResult

    Response = MsgBox( _
        "Do you want to export VBA modules?" & vbCrLf & vbCrLf & _
        "Destination:" & vbCrLf & _
        p_Destination & vbCrLf & vbCrLf & _
        "Click [Yes] to proceed or [No] to cancel.", _
        vbYesNo + vbQuestion, "Confirm VBA Export")

    ConfirmExport = (Response = vbYes)
End Function

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
    Dim importPath As String
    Dim fileName As String
    Dim ModuleName As String

    importPath = ThisWorkbook.Path & Application.PathSeparator & "src" & Application.PathSeparator

    If Not IsVBProjectAccessible() Then
        MsgBox GetVBATrustAccessMessage(), vbCritical, "VBA Project Access Error"
        Exit Sub
    End If

    ' Confirmation message
    UserResponse = MsgBox( _
        "Do you want to import VBA modules?" & vbCrLf & vbCrLf & _
        "Source Folder:" & vbCrLf & _
        importPath & vbCrLf & vbCrLf & _
        "Click [Yes] to proceed or [No] to cancel.", _
        vbYesNo + vbQuestion, "Confirm VBA Import")
    If UserResponse = vbNo Then
        Exit Sub
    End If
    
    ' Import each component
    fileName = Dir(importPath & "*")
    Do While fileName <> ""
        ModuleName = Left(fileName, InStrRev(fileName, ".") - 1)
        If Left(ModuleName, 2) = "wb" Or Left(ModuleName, 3) = "sht" Then
            ' TODO: Implement manual import logic here
        Else
            RemoveVBCompByName ModuleName
            ThisWorkbook.VBProject.VBComponents.Import importPath & fileName
        End If

        fileName = Dir
    Loop

    MsgBox "VBA code and Excel objects have been imported!", vbInformation
End Sub

Private Sub RemoveVBCompByName(vbCompName As String)
    Dim vbComp As Object
    
    On Error Resume Next
    Set vbComp = ThisWorkbook.VBProject.VBComponents(vbCompName)
    If Not vbComp Is Nothing Then
        ThisWorkbook.VBProject.VBComponents.Remove vbComp
    End If
    
    Err.Clear
    On Error GoTo 0
End Function

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


