VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTaskRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME As String = "clsTaskRepository"

Private m_objWorksheet As Worksheet

Public Sub Init(p_objWorksheet As Worksheet)
    Set m_objWorksheet = p_objWorksheet
End Sub

Public Function FindAll() As Collection
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "FindAll"

    Dim colTasks As Collection
    Set colTasks = New Collection
    
    Dim i As Long
    Dim LastRow As Long
    LastRow = m_objWorksheet.Cells(m_objWorksheet.Rows.Count, 1).End(xlUp).Row
    Dim objTask As clsTask
    
    For i = 2 To LastRow
        Set objTask = New clsTask

        With objTask
            .Id = CStr(m_objWorksheet.Cells(i, 1).Value)
            .WbsCode = CStr(m_objWorksheet.Cells(i, 2).Value)
            .Name = CStr(m_objWorksheet.Cells(i, 3).Value)
            .StartDate = ToDate(m_objWorksheet.Cells(i, 4).Value)
            .EndDate = ToDate(m_objWorksheet.Cells(i, 5).Value)
            .Duration = CDbl(Val(m_objWorksheet.Cells(i, 6).Value))
            .AssignedWorkHours = CDbl(Val(m_objWorksheet.Cells(i, 7).Value))
            .Progress = CDbl(Val(m_objWorksheet.Cells(i, 8).Value))
            .ActualStartDate = ToDate(m_objWorksheet.Cells(i, 9).Value)
            .ActualEndDate = ToDate(m_objWorksheet.Cells(i, 10).Value)
            .ActualWorkHours = CDbl(Val(m_objWorksheet.Cells(i, 11).Value))
            .BaselineStartDate = ToDate(m_objWorksheet.Cells(i, 12).Value)
            .BaselineEndDate = ToDate(m_objWorksheet.Cells(i, 13).Value)
            .BaselineWorkHours = CDbl(Val(m_objWorksheet.Cells(i, 14).Value))
            .PredecessorId = CStr(m_objWorksheet.Cells(i, 15).Value)
            .ParentId = CStr(m_objWorksheet.Cells(i, 16).Value)
            .ResourceId = CStr(m_objWorksheet.Cells(i, 17).Value)
        End With

        colTasks.Add objTask
    Next i

    Set FindAll = colTasks
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

'
' Return Nothing if task is not found
'
Public Function Find(p_Id As String) As clsTask
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "Find"

    Dim colTasks As Collection
    Set colTasks = Me.FindAll()

    Dim objTask As clsTask
    For Each objTask In colTasks
        If objTask.Id = p_Id Then
            Set Find = objTask
            Exit Function
        End If
    Next
    
    Set Find = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

Public Sub Save(p_Task As clsTask)
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "Save"

    Dim LastRow As Long
    LastRow = m_objWorksheet.Cells(m_objWorksheet.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim FoundRow As Long
    FoundRow = 0

    For i = 2 To LastRow
        If CStr(m_objWorksheet.Cells(i, 1).Value) = p_Task.Id Then
            FoundRow = i
            Exit For
        End If
    Next i

    Dim TargetRow As Long
    If FoundRow > 0 Then
        TargetRow = FoundRow
    Else
        TargetRow = LastRow + 1
    End If

    With m_objWorksheet
        .Cells(TargetRow, 1).Value = p_Task.Id
        .Cells(TargetRow, 2).Value = p_Task.WbsCode
        .Cells(TargetRow, 3).Value = p_Task.Name
        .Cells(TargetRow, 4).Value = p_Task.StartDate
        .Cells(TargetRow, 5).Value = p_Task.EndDate
        .Cells(TargetRow, 6).Value = p_Task.Duration
        .Cells(TargetRow, 7).Value = p_Task.AssignedWorkHours
        .Cells(TargetRow, 8).Value = p_Task.Progress
        .Cells(TargetRow, 9).Value = p_Task.ActualStartDate
        .Cells(TargetRow, 10).Value = p_Task.ActualEndDate
        .Cells(TargetRow, 11).Value = p_Task.ActualWorkHours
        .Cells(TargetRow, 12).Value = p_Task.BaselineStartDate
        .Cells(TargetRow, 13).Value = p_Task.BaselineEndDate
        .Cells(TargetRow, 14).Value = p_Task.BaselineWorkHours
        .Cells(TargetRow, 15).Value = p_Task.PredecessorId
        .Cells(TargetRow, 16).Value = p_Task.ParentId
        .Cells(TargetRow, 17).Value = p_Task.ResourceId
    End With

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Sub

Public Sub Delete(p_Id As String)
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "Delete"

    Dim LastRow As Long
    LastRow = m_objWorksheet.Cells(m_objWorksheet.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 2 To LastRow
        If CStr(m_objWorksheet.Cells(i, 1).Value) = p_Id Then
            m_objWorksheet.Rows(i).Delete
            Exit Sub
        End If
    Next i

    Err.Raise vbObjectError + 1000, MODULE_NAME & "." & PROCEDURE_NAME, "Task ID not fouind: " & p_Id
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Sub

'
' Check if the task with the given ID exists
'
Public Function Exists(p_Id As String) As Boolean
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "Exists"

    Exists = Not (Me.Find(p_Id) Is Nothing)
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

'
' Return the number of tasks
'
Public Function Count() As Long
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "Count"

    Count = Me.FindAll().Count
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

'
' Return 0 for not date value
'
Private Function ToDate(p_Value As Variant) As Date
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "ToDate"

    If IsDate(p_Value) Then
        ToDate = CDate(p_Value)
    Else
        ToDate = 0
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

