VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWbsApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_objTaskRepository As clsTaskRepository
Private m_objTaskService As clsTaskService
Private m_objResourceRepository As clsResourceRepository
Private m_objWbsPresenter As clsWbsPresenter
Private m_objWbsView As clsWbsView
Private m_objWbsController As clsWbsController

Public Sub Init()
    On Error GoTo ErrorHandler

    Set m_objTaskRepository = New clsTaskRepository
    Call m_objTaskRepository.Init(shtTasks)
    
    Set m_objTaskService = New clsTaskService
    Call m_objTaskService.Init(m_objTaskRepository)
    
    Set m_objResourceRepository = New clsResourceRepository
    Call m_objResourceRepository.Init(shtResources)
    
    Set m_objWbsPresenter = New clsWbsPresenter
    Call m_objWbsPresenter.Init(m_objTaskService)
    
    Set m_objWbsView = New clsWbsView
    Call m_objWbsView.Init(shtWbsView)
    
    Set m_objWbsController = New clsWbsController
    Call m_objWbsController.Init(m_objTaskRepository, _
                                 m_objTaskService, _
                                 m_objResourceRepository, _
                                 m_objWbsPresenter, _
                                 m_objWbsView)
    
    Exit Sub

ErrorHandler:
    MsgBox "Error on initializing application: " & Err.Description, vbCritical
End Sub

Public Sub CreateNewTask()
    On Error GoTo ErrorHandler
    
    Call m_objWbsController.CreateNewTask
    
    Exit Sub
ErrorHandler:
    MsgBox "Error on creating new task: " & Err.Description, vbCritical
End Sub

Public Sub RefreshWbsView()
    On Error GoTo ErrorHandler
    
    Call m_objWbsController.RefreshWbsView
    
    Exit Sub
ErrorHandler:
    MsgBox "Error on refreshing WBS view: " & Err.Description, vbCritical
End Sub

Public Sub HandleWbsCodeChanged(p_NewWbsCode As String, p_TaskId As String)
    On Error GoTo ErrorHandler

    Dim objTask As clsTask
    Set objTask = m_objTaskRepository.Find(p_TaskId)
    If objTask Is Nothing Then Exit Sub

    objTask.WbsCode = p_NewWbsCode
    
    Call m_objTaskRepository.Save(objTask)
'    Call m_objTaskService.RecalculateParentRelations
    Call m_objTaskService.RecalculateAll
    Call Me.RefreshWbsView

    Exit Sub
ErrorHandler:
    MsgBox "Error on updating WBS Code: " & Err.Description, vbCritical
End Sub
