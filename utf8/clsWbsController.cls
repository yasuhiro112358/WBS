VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWbsController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME As String = "clsWbsController"

Private m_objTaskRepository As clsTaskRepository
Private m_objTaskService As clsTaskService
Private m_objResourceRepository As clsResourceRepository
Private m_objWbsPresenter As clsWbsPresenter
Private m_objWbsView As clsWbsView

Public Sub Init(p_objTaskRepository As clsTaskRepository, _
                p_objTaskService As clsTaskService, _
                p_objResourceRepository As clsResourceRepository, _
                p_objWbsPresenter As clsWbsPresenter, _
                p_objWbsView As clsWbsView)
                
    Set m_objTaskRepository = p_objTaskRepository
    Set m_objTaskService = p_objTaskService
    Set m_objResourceRepository = p_objResourceRepository
    Set m_objWbsPresenter = p_objWbsPresenter
    Set m_objWbsView = p_objWbsView
End Sub

Sub CreateNewTask()
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "CreateNewTask"

    Dim objTask As clsTask
    Set objTask = New clsTask
    With objTask
        .Id = m_objTaskService.GenerateNewId()
        .WbsCode = m_objTaskService.GenerateInitialWbsCode()
        .Name = "New task"
        .BaselineStartDate = Date
        .BaselineEndDate = Date
        .BaselineWorkHours = 8
'        .AssignedWorkHours = 0
        .Progress = 0
'        .ParentId = ""
        .ResourceId = "R003"
    End With

    Call m_objTaskRepository.Save(objTask)

    Dim colTasks As Collection
    Set colTasks = m_objTaskRepository.FindAll()

    Dim colResources As Collection
    Set colResources = m_objResourceRepository.FindAll()

    Dim colViewItems As Collection
    Set colViewItems = m_objWbsPresenter.BuildWbsViewItems(colTasks, colResources)

    Call m_objWbsView.Clear
    Call m_objWbsView.Render(colViewItems)

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, PROCEDURE_NAME
End Sub

Sub DeleteTask(p_Id As String)
    ' Repositoryから削除
    ' 関連タスクの再計算（Service）
End Sub

Sub UpdateProgress(p_Id As String, p_Progress As Double)
    ' Repositoryで進捗を更新
    ' 親タスクの進捗も再計算（Service）
End Sub

Public Sub RefreshWbsView()
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "RefreshWBSView"

    Dim colTasks As Collection
    Set colTasks = m_objTaskRepository.FindAll()

    Dim colResources As Collection
    Set colResources = m_objResourceRepository.FindAll()

    Dim colViewItems As Collection
    Set colViewItems = m_objWbsPresenter.BuildWbsViewItems(colTasks, colResources)

    Call m_objWbsView.Clear
    Call m_objWbsView.Render(colViewItems)

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, PROCEDURE_NAME
End Sub

