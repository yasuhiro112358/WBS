Attribute VB_Name = "modTestTaskRepository"
Option Explicit

Sub TestTaskRepository()
    On Error GoTo ErrorHandler

    Dim objTaskRepository As clsTaskRepository
    Set objTaskRepository = New clsTaskRepository
    
    Dim colTasks As Collection
    Set colTasks = objTaskRepository.ReadAll()

    Debug.Print "== All tasks =="
    Dim objTask As clsTask
    For Each objTask In colTasks
'        Debug.Print "ID: " & objTask.Id & ", Name: " & objTask.Name & ", StartDate: " & objTask.StartDate
        Debug.Print "ID: " & objTask.Id, "Name: " & objTask.Name, "StartDate: " & objTask.StartDate
    Next

    Debug.Print "== Find by ID（e.g.;'1.1.1'） =="
    Dim objFound As clsTask
    Set objFound = objTaskRepository.FindById("1.1.1")

    If Not objFound Is Nothing Then
        Debug.Print "Found: " & objFound.Id & ", " & objFound.Name
    Else
        Debug.Print "Task is not found."
    End If

    Exit Sub

ErrorHandler:
    Call HandleError("modTestTaskRepository", "TestTaskRepository")
    
End Sub

