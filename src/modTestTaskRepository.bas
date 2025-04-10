Attribute VB_Name = "modTestTaskRepository"
Option Explicit

Sub TestTaskRepository()
    On Error GoTo ErrorHandler

    Dim objTaskRepository As clsTaskRepository
    Set objTaskRepository = New clsTaskRepository
    
    Dim colTasks As Collection
    Set colTasks = objTaskRepository.FindAll()

    Debug.Print "== All tasks =="
    Dim objTask As clsTask
    For Each objTask In colTasks
'        Debug.Print "ID: " & objTask.Id & ", Name: " & objTask.Name & ", StartDate: " & objTask.StartDate
        Debug.Print "ID: " & objTask.Id, "Name: " & objTask.Name, "StartDate: " & objTask.StartDate
    Next

    Debug.Print "== Find by ID（e.g.;'1.1.1'） =="
    Dim objFound As clsTask
    Set objFound = objTaskRepository.Find("1.1.1")

    If Not objFound Is Nothing Then
        Debug.Print "Found: " & objFound.Id & ", " & objFound.Name
    Else
        Debug.Print "Task is not found."
    End If

    Exit Sub

ErrorHandler:
    Call HandleError("modTestTaskRepository", "TestTaskRepository")
    
End Sub

Sub TestSaveTask()
    Dim objRepository As clsTaskRepository
    Set objRepository = New clsTaskRepository
    Dim objTask As clsTask
    Set objTask = New clsTask

    With objTask
        .Id = "1.1.2"
        .Name = "仕様書作成"
        .StartDate = #4/6/2025#
        .EndDate = #4/7/2025#
        .Duration = 2
        .AssignedWorkHours = 8
        .Progress = 0
        .ActualStartDate = 0
        .ActualEndDate = 0
        .ActualWorkHours = 0
        .BaselineStartDate = #4/6/2025#
        .BaselineEndDate = #4/7/2025#
        .BaselineWorkHours = 8
        .PredecessorId = "1.1.1"
        .ParentId = "1.1"
        .ResourceId = "R004"
    End With

    objRepository.Save objTask
    
    LogMessage "Complete saving."
End Sub

Sub TestDeleteTask()
    Dim objRepository As New clsTaskRepository
    objRepository.Delete "1.1.2"
    
    LogMessage "Complete deleting."
End Sub

Sub TestRepositoryExtras()
    On Error GoTo ErrorHandler

    Dim Repository As clsTaskRepository
    Set Repository = New clsTaskRepository

    ' === Exists ===
    Debug.Print "== Exists('1.1.1') = " & Repository.Exists("1.1.1")
    Debug.Print "== Exists('nonexistent') = " & Repository.Exists("nonexistent")

    ' === Count ===
    Debug.Print "== Count = " & Repository.Count()

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "TestRepositoryExtras"
End Sub


