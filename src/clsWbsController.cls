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
        .Progress = 0
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

Public Sub RefreshWbsView()
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "RefreshWBSView"

    Call m_objTaskService.RecalculateAll

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

Public Sub UpdateTask(ByVal p_objTarget As Range)
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
    Const COL_TASK_ID As Long = 18
    Const DATA_START_ROW As Long = 5
   
    If p_objTarget.Cells.Count > 1 Then
        Err.Raise vbObjectError + 1000, , "Changing multiple cells is not allowed."
    End If
    
    If p_objTarget.Row < DATA_START_ROW Then
        Exit Sub
    End If
    
    Dim TaskId As String
    TaskId = shtWbsView.Cells(p_objTarget.Row, COL_TASK_ID).Value
    If TaskId = "" Then
        Err.Raise vbObjectError + 1000, , "Task ID not found: " & Err.Description
    End If

    Dim objTask As clsTask
    Set objTask = m_objTaskRepository.Find(TaskId)
    If objTask Is Nothing Then
        Err.Raise vbObjectError + 1000, , "Task not found: " & TaskId
    End If
    
    Select Case p_objTarget.Column
        Case COL_WBS_CODE
            objTask.WbsCode = Trim(CStr(p_objTarget.Value))
        Case COL_TASK_NAME
            objTask.Name = Trim(CStr(p_objTarget.Value))
        Case COL_PROGRESS
            objTask.Progress = p_objTarget.Value
        Case COL_BASELINE_START_DATE
            objTask.BaselineStartDate = p_objTarget.Value
        Case COL_BASELINE_END_DATE
            objTask.BaselineEndDate = p_objTarget.Value
        Case COL_BASELINE_WORK_HOURS
            objTask.BaselineWorkHours = p_objTarget.Value
        Case COL_ASSIGNED_WORK_HOURS
            objTask.AssignedWorkHours = p_objTarget.Value
        Case COL_ACTUAL_WORK_HOURS
            objTask.ActualWorkHours = p_objTarget.Value
        Case COL_ACTUAL_START_DATE
            objTask.ActualStartDate = p_objTarget.Value
        Case COL_ACTUAL_END_DATE
            objTask.ActualEndDate = p_objTarget.Value
        Case COL_PREDECESSOR_WBS_Code
            objTask.PredecessorId = m_objTaskService.FindIdByWbsCode(Trim(CStr(p_objTarget.Value)))
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

