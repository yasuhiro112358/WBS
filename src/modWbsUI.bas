Attribute VB_Name = "modWbsUI"
Option Explicit

Public Sub BtnCreateNewTask()
    On Error GoTo ErrorHandler
    Application.EnableEvents = False
    
    Dim objWbsApplication As clsWbsApplication
    Set objWbsApplication = New clsWbsApplication
    Call objWbsApplication.Init
    Call objWbsApplication.CreateNewTask

Cleanup:
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Failed to add a new task: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub BtnRefreshWBSView()
    On Error GoTo ErrorHandler
    Application.EnableEvents = False
    
    Dim objWbsApplication As clsWbsApplication
    Set objWbsApplication = New clsWbsApplication
    Call objWbsApplication.Init
    Call objWbsApplication.RefreshWbsView
    
Cleanup:
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Failed to refresh WBS view: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

