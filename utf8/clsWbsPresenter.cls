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
    
    Dim dictTaskIdToLineNo As clsDictionary
    Set dictTaskIdToLineNo = New clsDictionary
    
    Dim dictTaskIdToTask As clsDictionary
    Set dictTaskIdToTask = New clsDictionary

    Dim dictWbsCodeToTask As clsDictionary
    Set dictWbsCodeToTask = New clsDictionary
    
    Dim LineNo As Long
    LineNo = 1
    Dim objTask As clsTask
    For Each objTask In p_colTasks
        Call dictTaskIdToLineNo.Add(objTask.Id, LineNo)
        
        Call dictTaskIdToTask.Add(objTask.Id, objTask)
        
        If objTask.WbsCode <> "" Then
            Call dictWbsCodeToTask.Add(objTask.WbsCode, objTask)
        End If
        LineNo = LineNo + 1
    Next


    Dim colWbsViewItems As Collection
    Set colWbsViewItems = New Collection
    ' Dim objTask As clsTask
    Dim objWbsViewItem As clsWbsViewItem
    
'     Dim objPredecessorTask As clsTask
    Dim objResource As clsResource
    
    LineNo = 1
    For Each objTask In p_colTasks
        Set objWbsViewItem = New clsWbsViewItem
        With objWbsViewItem
            .TaskId = objTask.Id
            .WbsCode = objTask.WbsCode
            .PredecessorId = objTask.PredecessorId
            .ResourceId = objTask.ResourceId
            
            .TaskLineNo = LineNo
            If dictTaskIdToLineNo.Exists(objTask.ParentId) Then
                .ParentLineNo = dictTaskIdToLineNo.Item(objTask.ParentId)
            End If
            If dictTaskIdToLineNo.Exists(objTask.PredecessorId) Then
                .PredecessorLineNo = dictTaskIdToLineNo.Item(objTask.PredecessorId)
            End If

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
            
            .StatusLabel = m_objTaskService.GetStatus(objTask)
            
'            If dictTaskIdToTask.Exists(objTask.PredecessorId) Then
'                .PredecessorName = dictTaskIdToTask.Item(objTask.PredecessorId).Name
'            End If
            If objTask.PredecessorId <> "" And dictTaskIdToTask.Exists(objTask.PredecessorId) Then
                Dim objPredecessorTask As clsTask
                Set objPredecessorTask = dictTaskIdToTask.Item(objTask.PredecessorId)
                .PredecessorWbsCode = objPredecessorTask.WbsCode
                .PredecessorName = objPredecessorTask.Name
            End If
            
        End With
        
'        Set objPredecessorTask = FindTaskById(p_colTasks, objTask.PredecessorId)
'        If Not objPredecessorTask Is Nothing Then
'            objWbsViewItem.PredecessorName = objPredecessorTask.Name
'        Else
'            objWbsViewItem.PredecessorName = ""
'        End If

        Set objResource = FindResourceById(p_colResources, objTask.ResourceId)
        If Not objResource Is Nothing Then
            objWbsViewItem.ResourceName = objResource.Name
        Else
            objWbsViewItem.ResourceName = "(Unassigned)"
        End If

        colWbsViewItems.Add objWbsViewItem
        
        LineNo = LineNo + 1
    Next objTask

'    Set BuildWbsViewItems = colWbsViewItems
    Set BuildWbsViewItems = SortCollectionByProperty(colWbsViewItems, "WbsCode")
    
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

Private Function CreateDisplayName(p_Name As String, p_Level As Long) As String
    CreateDisplayName = String((p_Level - 1) * INDENT_WIDTH, " ") & LINE_PREFIX & p_Name
End Function

Private Function FindTaskById(p_colTasks As Collection, p_Id As String) As clsTask
    Dim objTask As clsTask
    For Each objTask In p_colTasks
        If objTask.Id = p_Id Then
            Set FindTaskById = objTask
            Exit Function
        End If
    Next
    
    Set FindTaskById = Nothing
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
