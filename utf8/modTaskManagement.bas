Attribute VB_Name = "modTaskManagement"
Sub AddTask()
    Dim ws As Worksheet
    Set ws = shtWBS
    
    Dim LastRow As Long
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    'ws.Cells(lastRow, 1).value = lastRow - 1 ' Task ID
    ws.Cells(LastRow, 1).Value = GenerateUUIDv4()
    ws.Cells(LastRow, 2).Value = InputBox("Enter Task Name:")
    ws.Cells(LastRow, 7).Value = InputBox("Enter Baseline Start Date (yyyy/mm/dd):")
    ws.Cells(LastRow, 8).Value = InputBox("Enter Baseline End Date (yyyy/mm/dd):")
    ws.Cells(LastRow, 9).Value = InputBox("Enter Baseline Work Hours:")
    
    MsgBox "Task has been added successfully!", vbInformation
End Sub

Sub RefreshTask()
    Dim sht As Worksheet
    Set sht = shtWBS
    
    Dim colDict As clsCustomDictionary
    Set colDict = GetColDict()
    
    Dim Row As Integer
    Dim col As Integer
    Dim TaskId As String
    Row = InputBox("Enter row # to refresh:")
    col = colDict.Item("Task ID")
    TaskId = sht.Cells(Row, col).Value
    
    'various task will be included
    RefreshTaskById TaskId

    LogMessage "[RefreshTask] Completed"
End Sub

Sub RefreshTaskById(ByVal TaskId As String)
    Dim sht As Worksheet
    Set sht = shtWBS
    
    Dim taskRow As Integer
    taskRow = getTaskRow(TaskId)
    If taskRow = 0 Then
        Exit Sub
    End If
        
    Dim colDict As clsCustomDictionary
    Set colDict = GetColDict()
                
    
    ' Start Date
    Dim BaselineStartDate As Date
    Dim ActualStartDate As Date
    Dim StartDate As Date
    
    If IsDate(sht.Cells(taskRow, colDict.Item("Baseline Start Date")).Value) Then
        BaselineStartDate = sht.Cells(taskRow, colDict.Item("Baseline Start Date")).Value
    Else
        BaselineStartDate = 0
    End If
    
    If IsDate(sht.Cells(taskRow, colDict.Item("Actual Start Date")).Value) Then
        ActualStartDate = sht.Cells(taskRow, colDict.Item("Actual Start Date")).Value
    Else
        ActualStartDate = 0
    End If
    
    If ActualStartDate <> 0 Then
        StartDate = ActualStartDate
    Else
        StartDate = BaselineStartDate
    End If
    sht.Cells(taskRow, colDict.Item("Start Date")).Value = StartDate
    
    ' End Date
    Dim BaselineEndDate As Date
    Dim ActualEndDate As Date
    Dim EndDate As Date
    
    If IsDate(sht.Cells(taskRow, colDict.Item("Baseline End Date")).Value) Then
        BaselineEndDate = sht.Cells(taskRow, colDict.Item("Baseline End Date")).Value
    Else
        BaselineEndDate = 0
    End If
    
    If IsDate(sht.Cells(taskRow, colDict.Item("Actual End Date")).Value) Then
        ActualEndDate = sht.Cells(taskRow, colDict.Item("Actual End Date")).Value
    Else
        ActualEndDate = 0
    End If
    
    If ActualEndDate <> 0 Then
        EndDate = ActualEndDate
    Else
        EndDate = BaselineEndDate
    End If
    sht.Cells(taskRow, colDict.Item("End Date")).Value = EndDate



    
    ' Remaining Work Hours
    sht.Cells(taskRow, colDict.Item("Remaining Work Hours")).Value = sht.Cells(taskRow, colDict.Item("Assigned Work Hours")).Value - sht.Cells(taskRow, colDict.Item("Actual Work Hours")).Value
    ' Progress (%)
    sht.Cells(taskRow, colDict.Item("Progress (%)")).Value = (sht.Cells(taskRow, colDict.Item("Actual Work Hours")).Value / sht.Cells(taskRow, colDict.Item("Assigned Work Hours")).Value) * 100

    LogMessage "[RefreshTaskById] Task ID: " & TaskId & " has been updated"
End Sub

'
' Return 0 if `TaskId` is not found
'
Function getTaskRow(TaskId As String) As Integer
    Dim sht As Worksheet
    Set sht = shtWBS
   
    Dim LastRow As Integer
    LastRow = sht.Cells(sht.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Integer
    For i = 2 To LastRow
        If sht.Cells(i, 1).Value = TaskId Then
            getTaskRow = i
            Exit Function
        End If
    Next i
    
    getTaskRow = 0
End Function

