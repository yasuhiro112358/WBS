VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWbsPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME As String = "clsWbsPresenter"

Private Const INDENT_WIDTH As Long = 8
Private Const LINE_PREFIX As String = "* "

Private m_objTaskService As clsTaskService

Public Sub Init(p_objTaskService As clsTaskService)
    Set m_objTaskService = p_objTaskService
End Sub

Public Function BuildWbsViewItems(p_colTasks As Collection, p_colResources As Collection) As Collection
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "BuildViewItems"

    Dim colWbsViewItems As Collection
    Set colWbsViewItems = New Collection

    Dim objTask As clsTask
    Dim objWbsViewItem As clsWbsViewItem
    Dim objResource As clsResource
    
    For Each objTask In p_colTasks
        Set objWbsViewItem = New clsWbsViewItem
        With objWbsViewItem
            .Id = objTask.Id
            .Name = objTask.Name
            
            .Level = m_objTaskService.GetLevel(objTask)
            
            .DisplayName = CreateDisplayName(.Name, .Level)
            
            .StartDate = objTask.StartDate
            .EndDate = objTask.EndDate
            .Progress = objTask.Progress
            .Duration = objTask.Duration
            .BaselineStartDate = objTask.BaselineStartDate
            .BaselineEndDate = objTask.BaselineEndDate
            .BaselineWorkHours = objTask.BaselineWorkHours
            .AssignedWorkHours = objTask.AssignedWorkHours
            .ActualWorkHours = objTask.ActualWorkHours
            
            .RemainingWorkHours = m_objTaskService.CalcRemainingWorkHours(objTask)
            
            .ActualStartDate = objTask.ActualStartDate
            .ActualEndDate = objTask.ActualEndDate
            .PredecessorId = objTask.PredecessorId
            
            .StatusLabel = m_objTaskService.GetStatus(objTask)
        End With

        Set objResource = FindResourceById(p_colResources, objTask.ResourceId)
        If Not objResource Is Nothing Then
            objWbsViewItem.ResourceName = objResource.Name
        Else
            objWbsViewItem.ResourceName = "(Unassigned)"
        End If

        colWbsViewItems.Add objWbsViewItem
    Next

    Set BuildWbsViewItems = colWbsViewItems
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

Private Function CreateDisplayName(p_Name As String, p_Level As Long) As String
    CreateDisplayName = String((p_Level - 1) * INDENT_WIDTH, " ") & LINE_PREFIX & p_Name
End Function

Private Function FindResourceById(p_colResources As Collection, p_Id As String) As clsResource
    Dim objResource As clsResource
    For Each objResource In p_colResources
        If objResource.Id = p_Id Then
            Set FindResourceById = objResource
            Exit Function
        End If
    Next
    
    Set FindResourceById = Nothing
End Function
