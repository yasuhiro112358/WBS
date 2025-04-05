Attribute VB_Name = "modWbsUI"
Option Explicit

Public Sub BtnCreateNewTask()
    On Error GoTo ErrorHandler
    
    Dim objWbsApplication As clsWbsApplication
    Set objWbsApplication = New clsWbsApplication
    Call objWbsApplication.Init
    Call objWbsApplication.CreateNewTask
    
    Exit Sub
ErrorHandler:
    MsgBox "Failed to add a new task: " & Err.Description, vbCritical
End Sub

Public Sub BtnRefreshWBSView()
    On Error GoTo ErrorHandler
    
    Dim objWbsApplication As clsWbsApplication
    Set objWbsApplication = New clsWbsApplication
    Call objWbsApplication.Init
    Call objWbsApplication.RefreshWbsView
    
    Exit Sub
ErrorHandler:
    MsgBox "Failed to refresh WBS view: " & Err.Description, vbCritical
End Sub

