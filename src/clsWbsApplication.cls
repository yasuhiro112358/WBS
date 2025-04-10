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

Public Sub UpdateTask(p_Data As Variant, p_TaskId As String, p_DataColumun As Long)
    On Error GoTo ErrorHandler
    
    Const COL_WBS_CODE As Long = 1
    Const COL_TASK_NAME As Long = 2
    Const COL_PROGRESS As Long = 5
    Const COL_BASELINE_START_DATE As Long = 7
    Const COL_BASELINE_END_DATE As Long = 8
    Const COL_BASELINE_WORK_HOURS As Long = 9
    Const COL_ASSIGNED_WORK_HOURS As Long = 10
    Const COL_ACTUAL_WORK_HOURS As Long = 11
    Const COL_ACTUAL_START_DATE As Long = 13
    Const COL_ACTUAL_END_DATE As Long = 14
    Const COL_PREDECESSOR_WBS_Code As Long = 15
        
    Dim objTask As clsTask
    Set objTask = m_objTaskRepository.Find(p_TaskId)
    If objTask Is Nothing Then
        Err.Raise vbObjectError + 1000, , "Task not found: " & p_TaskId
    End If
    
    Select Case p_DataColumun
        Case COL_WBS_CODE
            objTask.WbsCode = Trim(CStr(p_Data))
        Case COL_TASK_NAME
            objTask.Name = Trim(CStr(p_Data))
        Case COL_PROGRESS
            objTask.Progress = p_Data
        Case COL_BASELINE_START_DATE
            objTask.BaselineStartDate = p_Data
        Case COL_BASELINE_END_DATE
            objTask.BaselineEndDate = p_Data
        Case COL_BASELINE_WORK_HOURS
            objTask.BaselineWorkHours = p_Data
        Case COL_ASSIGNED_WORK_HOURS
            objTask.AssignedWorkHours = p_Data
        Case COL_ACTUAL_WORK_HOURS
            objTask.ActualWorkHours = p_Data
        Case COL_ACTUAL_START_DATE
            objTask.ActualStartDate = p_Data
        Case COL_ACTUAL_END_DATE
            objTask.ActualEndDate = p_Data
        Case COL_PREDECESSOR_WBS_Code
            objTask.PredecessorId = m_objTaskService.FindIdByWbsCode(Trim(CStr(p_Data)))
        Case Else
            Err.Raise vbObjectError + 1000, , "Not allowed to change the data on this column."
    End Select
    
    ' Want to add validation here
    
    Call m_objTaskRepository.Save(objTask)
    
    Call m_objTaskService.RecalculateAll

    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
