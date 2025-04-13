VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTaskService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME As String = "clsTaskService"

Private m_objTaskRepository As clsTaskRepository

Public Sub Init(p_objTaskRepository As clsTaskRepository)
    Set m_objTaskRepository = p_objTaskRepository
End Sub

Public Function GenerateNewId() As String
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "GenerateNewId"

    GenerateNewId = GenerateUUIDv4()

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

Public Function GenerateInitialWbsCode() As String
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "GenerateInitialWbsCode"

    Const PREFIX As String = "WBS-"

    Dim colTasks As Collection
    Set colTasks = m_objTaskRepository.FindAll()

    Dim MaxCode As Long
    MaxCode = 0

    Dim objTask As clsTask
    For Each objTask In colTasks
        Dim Code As String
        Code = Trim(objTask.WbsCode)
        Code = Mid(Code, Len(PREFIX) + 1)
        
        If Code = "" Then
            GoTo Continue
        End If
        
        If InStr(Code, ".") <> 0 Then
            GoTo Continue
        End If
        
        If IsNumeric(Code) = False Then
            GoTo Continue
        End If
        
        If CLng(Code) > MaxCode Then
            MaxCode = CLng(Code)
        End If
            
Continue:
    Next objTask

    GenerateInitialWbsCode = CStr(PREFIX & (MaxCode + 1))
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

'
' Return actual if it exists, otherwise, return baseline
'
Public Function CalcStartDate(p_Task As clsTask) As Date
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "CalcStartDate"

    If p_Task.ActualStartDate <> 0 Then
        CalcStartDate = p_Task.ActualStartDate
    ElseIf p_Task.BaselineStartDate <> 0 Then
        CalcStartDate = p_Task.BaselineStartDate
    Else
        CalcStartDate = 0
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

Public Function CalcEndDate(p_Task As clsTask) As Date
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "CalcEndDate"

    If p_Task.ActualEndDate <> 0 Then
        CalcEndDate = p_Task.ActualEndDate
    ElseIf p_Task.BaselineEndDate <> 0 Then
        CalcEndDate = p_Task.BaselineEndDate
    Else
        CalcEndDate = 0
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

'
' If Progress exists remain it, if not calculate it
'
Public Function CalcProgress(p_Task As clsTask) As Double
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "CalcProgress"

    If p_Task.Progress > 0 Then
        CalcProgress = p_Task.Progress
        Exit Function
    End If

    If p_Task.AssignedWorkHours > 0 Then
        CalcProgress = (p_Task.ActualWorkHours / p_Task.AssignedWorkHours) * 100
        If CalcProgress > 100 Then CalcProgress = 100
    Else
        CalcProgress = 0
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

'
' Calculate difference between actual times, if it does not exists, calculate between baseline times
'
Public Function CalcDuration(p_Task As clsTask) As Double
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "CalcDuration"

    Dim StartDate As Date
    Dim EndDate As Date

    If p_Task.ActualStartDate <> 0 And p_Task.ActualEndDate <> 0 Then
        StartDate = p_Task.ActualStartDate
        EndDate = p_Task.ActualEndDate
    ElseIf p_Task.BaselineStartDate <> 0 And p_Task.BaselineEndDate <> 0 Then
        StartDate = p_Task.BaselineStartDate
        EndDate = p_Task.BaselineEndDate
    Else
        CalcDuration = 0
        Exit Function
    End If

    If EndDate < StartDate Then
        CalcDuration = 0
    Else
        CalcDuration = EndDate - StartDate + 1
    End If

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

'
' Return `Assigned - Actual`
'
Public Function CalcRemainingWorkHours(p_objTask As clsTask) As Double
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "CalcRemainingWorkHours"

    CalcRemainingWorkHours = p_objTask.AssignedWorkHours - p_objTask.ActualWorkHours
    If CalcRemainingWorkHours < 0 Then
        CalcRemainingWorkHours = 0
    End If

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

Public Sub RecalculateAll()
    Call Me.RecalculateParentRelations
    
    Dim colTasks As Collection
    Set colTasks = m_objTaskRepository.FindAll()
    
    Dim objTask As clsTask
    For Each objTask In colTasks
        objTask.StartDate = CalcStartDate(objTask)
        objTask.EndDate = CalcEndDate(objTask)
        objTask.Duration = CalcDuration(objTask)
        objTask.Progress = CalcProgress(objTask)
        
        Call m_objTaskRepository.Save(objTask)
    Next
    For Each objTask In colTasks
        Call Me.AdjustChildrenBaselineDates(objTask)
        Call AdjustBaselineStartDateByPredecessor(objTask)
    Next
End Sub

Public Function GetLevel(p_objTask As clsTask) As Long
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "GetLevel"

    Dim Level As Long
    Level = 1
    
    Dim objCurrentTask As clsTask
    Set objCurrentTask = p_objTask
    
    Do While Len(objCurrentTask.ParentId) > 0
        Set objCurrentTask = m_objTaskRepository.Find(objCurrentTask.ParentId)
        If objCurrentTask Is Nothing Then
            Exit Do
        End If
        Level = Level + 1
    Loop

    GetLevel = Level
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

'
' 親タスクの進捗を再計算（子タスクの平均など）
'
Public Sub RecalculateProgress(p_ParentId As String)
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "RecalculateProgress"

    ' TODO: 実装
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Sub

Public Function GetStatus(p_Task As clsTask) As String
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "GetStatus"

    If Me.HasUnfinishedPredecessors(p_Task) Then
        GetStatus = "Blocked"
    ElseIf p_Task.Progress >= 100 Or IsDate(p_Task.ActualEndDate) Then
        GetStatus = "Done"
    ElseIf p_Task.Progress > 0 Or IsDate(p_Task.ActualStartDate) Then
        GetStatus = "InProgress"
    Else
        GetStatus = "NotStarted"
    End If

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

Public Function HasUnfinishedPredecessors(p_Task As clsTask) As Boolean
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "HasUnfinishedPredecessors"

    Dim PredId As String
    PredId = Trim(p_Task.PredecessorId)
    
    If PredId = "" Then
        HasUnfinishedPredecessors = False
        Exit Function
    End If

    Dim objpredTask As clsTask
    Set objpredTask = m_objTaskRepository.Find(PredId)

    If objpredTask Is Nothing Then
        HasUnfinishedPredecessors = False
    ElseIf objpredTask.Progress < 100 Then
        HasUnfinishedPredecessors = True
    Else
        HasUnfinishedPredecessors = False
    End If

    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

Public Sub RecalculateParentRelations()
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "RecalculateParentRelations"

    Dim colTasks As Collection
    Set colTasks = m_objTaskRepository.FindAll()

    Dim dictWbsCodeToTaskId As clsDictionary
    Set dictWbsCodeToTaskId = New clsDictionary

    Dim objTask As clsTask
    For Each objTask In colTasks
        If Len(objTask.WbsCode) > 0 Then
            Call dictWbsCodeToTaskId.Add(objTask.WbsCode, objTask.Id)
        End If
    Next
    For Each objTask In colTasks
        Dim ParentWbsCode As String
        ParentWbsCode = GetParentWbsCode(objTask.WbsCode)

        If ParentWbsCode <> "" And _
           dictWbsCodeToTaskId.Exists(ParentWbsCode) Then
            objTask.ParentId = dictWbsCodeToTaskId.Item(ParentWbsCode)
        Else
            objTask.ParentId = ""
        End If

        Call m_objTaskRepository.Save(objTask)
    Next

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Sub

Private Function GetParentWbsCode(p_WbsCode As String) As String
    If InStrRev(p_WbsCode, ".") > 0 Then
        GetParentWbsCode = Left(p_WbsCode, InStrRev(p_WbsCode, ".") - 1)
    Else
        GetParentWbsCode = ""
    End If
End Function

Public Function FindIdByWbsCode(p_WbsCode As String) As String
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "FindIdByWbsCode"

    Dim colTasks As Collection
    Set colTasks = m_objTaskRepository.FindAll()
        
    Dim objTask As clsTask
    For Each objTask In colTasks
        If Trim(objTask.WbsCode) = Trim(p_WbsCode) Then
            FindIdByWbsCode = objTask.Id
            Exit Function
        End If
    Next
    
    FindIdByWbsCode = ""
    
'    Err.Raise vbObjectError + 1000, MODULE_NAME & "." & PROCEDURE_NAME, "Task with the WBS code not found."
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

Public Sub AdjustChildrenBaselineDates(p_objParent As clsTask)
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "AdjustChildrenBaselineDates"

    Dim colChildren As Collection
    Set colChildren = Me.FindChildren(p_objParent.Id)

    Dim objChild As clsTask
    For Each objChild In colChildren
        Dim Updated As Boolean
        Updated = False

        If objChild.BaselineStartDate = 0 Or _
           objChild.BaselineStartDate < p_objParent.BaselineStartDate Then
            objChild.BaselineStartDate = p_objParent.BaselineStartDate
            Updated = True
        End If

        If objChild.BaselineEndDate = 0 Or _
           objChild.BaselineEndDate > p_objParent.BaselineEndDate Then
            objChild.BaselineEndDate = p_objParent.BaselineEndDate
            Updated = True
        End If

        If Updated Then
            Call m_objTaskRepository.Save(objChild)
        End If
        
        Call AdjustChildrenBaselineDates(objChild)
    Next

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Sub

Public Function FindChildren(p_Id As String) As Collection
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "FindChildren"

    Dim colTasks As Collection
    Set colTasks = m_objTaskRepository.FindAll()

    Dim colChildren As Collection
    Set colChildren = New Collection

    Dim objTask As clsTask
    For Each objTask In colTasks
        If objTask.ParentId = p_Id Then
            colChildren.Add objTask
        End If
    Next

    Set FindChildren = colChildren
   
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

Private Sub AdjustBaselineStartDateByPredecessor(p_objTask As clsTask)
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "AdjustBaselineStartDateByPredecessor"

    Dim PredId As String
    PredId = Trim(p_objTask.PredecessorId)
    If PredId = "" Then
        Exit Sub
    End If

    Dim objPred As clsTask
    Set objPred = m_objTaskRepository.Find(PredId)
    If objPred Is Nothing Then
        Exit Sub
    End If
    
    Dim PredEnd As Date
    PredEnd = objPred.BaselineEndDate
    If PredEnd = 0 Then
        Exit Sub
    End If
    
    If p_objTask.BaselineStartDate < (PredEnd + 1) Then
        p_objTask.BaselineStartDate = PredEnd + 1
        p_objTask.BaselineEndDate = p_objTask.BaselineStartDate + p_objTask.Duration - 1
        Call m_objTaskRepository.Save(p_objTask)
    End If

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Sub
