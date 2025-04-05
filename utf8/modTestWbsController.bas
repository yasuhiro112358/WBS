Attribute VB_Name = "modTestWbsController"
Option Explicit

Sub TestCreateNewTask()
    On Error GoTo ErrorHandler

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
    
    Call objWbsController.CreateNewTask

    MsgBox "Creating new Task, showed it onto WBS view.", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Error" & Err.Description, vbCritical
End Sub


Sub TestRefreshWbsView()
    On Error GoTo ErrorHandler

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
    
    Call objWbsController.RefreshWbsView

    MsgBox "Refreshed WBS view.", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Error" & Err.Description, vbCritical
End Sub



