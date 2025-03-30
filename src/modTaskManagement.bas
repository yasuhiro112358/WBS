Attribute VB_Name = "modTaskManagement"
Sub AddTask()
    Dim ws As Worksheet
    Set ws = shtWBS
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    
    'ws.Cells(lastRow, 1).value = lastRow - 1 ' Task ID
    ws.Cells(lastRow, 1).value = GenerateUUIDv4()
    ws.Cells(lastRow, 2).value = InputBox("Enter Task Name:")
    ws.Cells(lastRow, 7).value = InputBox("Enter Baseline Start Date (yyyy/mm/dd):")
    ws.Cells(lastRow, 8).value = InputBox("Enter Baseline End Date (yyyy/mm/dd):")
    ws.Cells(lastRow, 9).value = InputBox("Enter Baseline Work Hours:")
    
    MsgBox "Task has been added successfully!", vbInformation
End Sub

Sub RefreshTask()
    Dim sht As Worksheet
    Set sht = shtWBS
    
    Dim colDict As clsCustomDictionary
    Set colDict = GetColDict()
    
    Dim row As Integer
    Dim col As Integer
    Dim taskId As String
    row = InputBox("Enter row # to refresh:")
    col = colDict.Item("Task ID")
    taskId = sht.Cells(row, col).value
    
    'various task will be included
    RefreshTaskById taskId

    LogMessage "[RefreshTask] Completed"
End Sub

Sub RefreshTaskById(ByVal taskId As String)
    Dim sht As Worksheet
    Set sht = shtWBS
    
    Dim taskRow As Integer
    taskRow = getTaskRow(taskId)
    If taskRow = 0 Then
        Exit Sub
    End If
        
    Dim colDict As clsCustomDictionary
    Set colDict = GetColDict()
                
    
    ' Start Date
    Dim baselineStartDate As Date
    Dim actualStartDate As Date
    Dim StartDate As Date
    
    If IsDate(sht.Cells(taskRow, colDict.Item("Baseline Start Date")).value) Then
        baselineStartDate = sht.Cells(taskRow, colDict.Item("Baseline Start Date")).value
    Else
        baselineStartDate = 0
    End If
    
    If IsDate(sht.Cells(taskRow, colDict.Item("Actual Start Date")).value) Then
        actualStartDate = sht.Cells(taskRow, colDict.Item("Actual Start Date")).value
    Else
        actualStartDate = 0
    End If
    
    If actualStartDate <> 0 Then
        StartDate = actualStartDate
    Else
        StartDate = baselineStartDate
    End If
    sht.Cells(taskRow, colDict.Item("Start Date")).value = StartDate
    
    ' End Date
    Dim baselineEndDate As Date
    Dim actualEndDate As Date
    Dim EndDate As Date
    
    If IsDate(sht.Cells(taskRow, colDict.Item("Baseline End Date")).value) Then
        baselineEndDate = sht.Cells(taskRow, colDict.Item("Baseline End Date")).value
    Else
        baselineEndDate = 0
    End If
    
    If IsDate(sht.Cells(taskRow, colDict.Item("Actual End Date")).value) Then
        actualEndDate = sht.Cells(taskRow, colDict.Item("Actual End Date")).value
    Else
        actualEndDate = 0
    End If
    
    If actualEndDate <> 0 Then
        EndDate = actualEndDate
    Else
        EndDate = baselineEndDate
    End If
    sht.Cells(taskRow, colDict.Item("End Date")).value = EndDate



    
    ' Remaining Work Hours
    sht.Cells(taskRow, colDict.Item("Remaining Work Hours")).value = sht.Cells(taskRow, colDict.Item("Assigned Work Hours")).value - sht.Cells(taskRow, colDict.Item("Actual Work Hours")).value
    ' Progress (%)
    sht.Cells(taskRow, colDict.Item("Progress (%)")).value = (sht.Cells(taskRow, colDict.Item("Actual Work Hours")).value / sht.Cells(taskRow, colDict.Item("Assigned Work Hours")).value) * 100

    LogMessage "[RefreshTaskById] Task ID: " & taskId & " has been updated"
End Sub

'
' Return 0 if `TaskId` is not found
'
Function getTaskRow(taskId As String) As Integer
    Dim sht As Worksheet
    Set sht = shtWBS
   
    Dim lastRow As Integer
    lastRow = sht.Cells(sht.Rows.Count, 1).End(xlUp).row
    
    Dim i As Integer
    For i = 2 To lastRow
        If sht.Cells(i, 1).value = taskId Then
            getTaskRow = i
            Exit Function
        End If
    Next i
    
    getTaskRow = 0
End Function

