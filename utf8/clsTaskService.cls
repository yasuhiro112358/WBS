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

'
' タスクの依存関係をチェック（前提タスクが未完了か？）
'
Public Function HasUnfinishedPredecessors(p_Task As clsTask) As Boolean
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "HasUnfinishedPredecessors"

    ' TODO: 実装
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

'
' タスクの状態を判定（例：NotStarted / InProgress / Done / Blocked）
'
Public Function GetStatus(p_Task As clsTask) As String
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "GetStatus"

    ' TODO: 実装（例：依存タスクが未完了なら "Blocked" など）
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

