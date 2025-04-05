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

    Dim colTasks As Collection
    Set colTasks = m_objTaskRepository.FindAll()

    Dim MaxCode As Long
    MaxCode = 0

    Dim objTask As clsTask
    For Each objTask In colTasks
        Dim Code As String
        Code = Trim(objTask.WbsCode)

'        If Code <> "" And InStr(Code, ".") = 0 Then
'            If IsNumeric(Code) Then
'                If CLng(Code) > MaxCode Then
'                    MaxCode = CLng(Code)
'                End If
'            End If
'        End If
        
        If Code = "" Then GoTo Continue
        If InStr(Code, ".") <> 0 Then GoTo Continue
        If IsNumeric(Code) = False Then GoTo Continue
        If CLng(Code) > MaxCode Then
            MaxCode = CLng(Code)
        End If
            
Continue:
    Next objTask

    GenerateInitialWbsCode = CStr(MaxCode + 1)
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
' 子タスク一覧を取得
'
Public Function FindChildren(p_ParentId As String) As Collection
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "FindChildren"

    ' TODO: 実装
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

        If ParentWbsCode <> "" And dictWbsCodeToTaskId.Exists(ParentWbsCode) Then
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


