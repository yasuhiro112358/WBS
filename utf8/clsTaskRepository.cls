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

Private m_objWorksheet As Worksheet

Private Sub Class_Initialize()
    Set m_objWorksheet = shtTasks
End Sub

Public Function ReadAll() As Collection
    On Error GoTo ErrorHandler

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
            .Name = CStr(m_objWorksheet.Cells(i, 2).Value)
            .StartDate = ToDate(m_objWorksheet.Cells(i, 3).Value)
            .EndDate = ToDate(m_objWorksheet.Cells(i, 4).Value)
            .Duration = CDbl(Val(m_objWorksheet.Cells(i, 5).Value))
            .AssignedWorkHours = CDbl(Val(m_objWorksheet.Cells(i, 6).Value))
            .Progress = CDbl(Val(m_objWorksheet.Cells(i, 7).Value))
            .ActualStartDate = ToDate(m_objWorksheet.Cells(i, 8).Value)
            .ActualEndDate = ToDate(m_objWorksheet.Cells(i, 9).Value)
            .ActualWorkHours = CDbl(Val(m_objWorksheet.Cells(i, 10).Value))
            .BaselineStartDate = ToDate(m_objWorksheet.Cells(i, 11).Value)
            .BaselineEndDate = ToDate(m_objWorksheet.Cells(i, 12).Value)
            .BaselineWorkHours = CDbl(Val(m_objWorksheet.Cells(i, 13).Value))
            .PredecessorId = CStr(m_objWorksheet.Cells(i, 14).Value)
            .ParentId = CStr(m_objWorksheet.Cells(i, 15).Value)
            .ResourceId = CStr(m_objWorksheet.Cells(i, 16).Value)
        End With

        colTasks.Add objTask
    Next i

    Set ReadAll = colTasks
    
    Exit Function
    
ErrorHandler:
    Call HandleError("clsTaskRepository", "ReadAll")

End Function

'
' Return Nothing if task is not found
'
Public Function FindById(p_Id As String) As clsTask
    On Error GoTo ErrorHandler

    Dim colTasks As Collection
    Set colTasks = Me.ReadAll()

    Dim objTask As clsTask
    For Each objTask In colTasks
        If objTask.Id = p_Id Then
            Set FindById = objTask
            Exit Function
        End If
    Next

    Set FindById = Nothing
    
    Exit Function
    
ErrorHandler:
    Call HandleError("clsTaskRepository", "FindById")

End Function

'
' Return 0 for not date value
'
Private Function ToDate(p_Value As Variant) As Date
    On Error GoTo ErrorHandler

    If IsDate(p_Value) Then
        ToDate = CDate(p_Value)
    Else
        ToDate = 0
    End If
    
    Exit Function
    
ErrorHandler:
    Call HandleError("clsTaskRepository", "ToDate")

End Function

