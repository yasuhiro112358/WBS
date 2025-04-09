Attribute VB_Name = "modTests"
Option Explicit

Sub TestGeneralError()
    On Error GoTo ErrorHandler
    
    Dim i As Long
    i = 10 / 0
    
    Dim j As Long
    j = "apple"
    
    Exit Sub
    
ErrorHandler:
    Call HandleError("modTests", "TestGeneralError")

End Sub

Sub TestGenerateUUIDv4()
   Dim uuid As String
   uuid = GenerateUUIDv4()
   Debug.Print uuid
End Sub

Sub TestCustomDictionary()
    Debug.Print "Start TestCustomDictionary"
    
    Dim Dict As clsCustomDictionary
    Set Dict = CustomDictionary()
    
    Dim keysArray As Variant
    Dim valuesArray As Variant
        
    ' Add Tasks
    Dict.Add "Task1", "Write VBA Code"
    Dict.Add "Task2", "Review Code"
    Dict.Item("Task3") = "Test Code"
    
    ' Get all keys
    keysArray = Dict.Keys()
    Dim keyArray As Variant
    For Each keyArray In keysArray
        Debug.Print "keyArray: " & keyArray
    Next
    
    ' Get all values
    valuesArray = Dict.Values()
    Dim valueArray As Variant
    For Each valueArray In valuesArray
        Debug.Print "valueArray: " & valueArray
    Next

    Debug.Print "Check if a key exist or not"
    Debug.Print "dict.Exists('Task2'): " & Dict.Exists("Task2")
    Debug.Print "dict.Exists('Unreal Task'): " & Dict.Exists("Unreal Task")

    Dim i As Integer
    For i = LBound(keysArray) To UBound(keysArray)
        Debug.Print keysArray(i) & " -> " & valuesArray(i)
    Next i

    Debug.Print "dict.Count: " & Dict.Count

    Dict.Remove "Task1"
    
    Debug.Print "dict.Count: " & Dict.Count

    Dict.RemoveAll
    
    Debug.Print "dict.Count: " & Dict.Count
End Sub

Sub TestGetColDict()
    Dim colDict As clsCustomDictionary
    Set colDict = GetColDict()
    
    Debug.Print colDict.Item("Task ID")
    Debug.Print colDict.Item("Baseline Start Date")
End Sub

Sub TestTask()
    Dim objTask As clsTask
    Set objTask = New clsTask
    
    objTask.Initialize "Sample Task 1", Date, "2025-4-1"
    
    Debug.Print objTask.Id
    Debug.Print objTask.Name
    Debug.Print objTask.StartDate
    Debug.Print objTask.EndDate
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("test")
    objTask.WriteToSheet ws, 2
    
    objTask.Name = objTask.Name & " updated"
    objTask.WriteToSheet ws, 3
        
     Set objTask = Nothing
End Sub

Sub TestTaskCol()
    Debug.Print TASK_START_DATE
    Debug.Print TASK_END_DATE
    Debug.Print TASK_BASELINE_START_DATE
End Sub

Sub TestA1Notation()
    Debug.Print COL_A__
    Debug.Print COL_ID_
    Debug.Print COL_IV_
End Sub

Sub TestGlobals()
    Debug.Print APP_NAME
End Sub

