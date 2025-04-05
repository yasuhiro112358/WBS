Attribute VB_Name = "modWbsUI"
Option Explicit

Private Function CreateWbsController() As clsWbsController
    Dim objTaskRepository As clsTaskRepository
    Set objTaskRepository = New clsTaskRepository
    
    Dim objTaskService As clsTaskService
    Set objTaskService = New clsTaskService
    Call objTaskService.Init(objTaskRepository)
    
    Dim objResourceRepository As clsResourceRepository
    Set objResourceRepository = New clsResourceRepository
    
    Dim objWbsPresenter As clsWbsPresenter
    Set objWbsPresenter = New clsWbsPresenter
    Call objWbsPresenter.Init(objTaskService)
    
    Dim objWbsView As clsWbsView
    Set objWbsView = New clsWbsView
    Call objWbsView.Init(shtWbsView)
    
    Dim objWbsController As clsWbsController
    Set objWbsController = New clsWbsController
    Call objWbsController.Init(objTaskRepository, objTaskService, objResourceRepository, objWbsPresenter, objWbsView)

    Set CreateWbsController = objWbsController
End Function

Public Sub BtnCreateNewTask()
    On Error GoTo ErrorHandler
    
    Dim objWbsController As clsWbsController
    Set objWbsController = CreateWbsController()

    Call objWbsController.CreateNewTask
    
    Exit Sub

ErrorHandler:
    MsgBox "Failed to add a new task: " & Err.Description, vbCritical
End Sub

Public Sub BtnRefreshWBSView()
    On Error GoTo ErrorHandler
    
    Dim objWbsController As clsWbsController
    Set objWbsController = CreateWbsController()
    Call objWbsController.RefreshWBSView
    
    Exit Sub

ErrorHandler:
    MsgBox "Failed to refresh WBS view: " & Err.Description, vbCritical
End Sub

