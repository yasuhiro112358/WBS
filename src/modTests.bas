Attribute VB_Name = "modTests"
Option Explicit

Sub TestGenerateUUIDv4()
   Dim uuid As String
   uuid = GenerateUUIDv4()
   Debug.Print uuid
End Sub

Sub TestCustomDictionary()
    Debug.Print "Start TestCustomDictionary"
    
    Dim dict As clsCustomDictionary
    Set dict = CustomDictionary()
    
    Dim keysArray As Variant
    Dim valuesArray As Variant
        
    ' Add Tasks
    dict.Add "Task1", "Write VBA Code"
    dict.Add "Task2", "Review Code"
    dict.Item("Task3") = "Test Code"
    
    ' Get all keys
    keysArray = dict.Keys()
    Dim keyArray As Variant
    For Each keyArray In keysArray
        Debug.Print "keyArray: " & keyArray
    Next
    
    ' Get all values
    valuesArray = dict.Values()
    Dim valueArray As Variant
    For Each valueArray In valuesArray
        Debug.Print "valueArray: " & valueArray
    Next

    Debug.Print "Check if a key exist or not"
    Debug.Print "dict.Exists('Task2'): " & dict.Exists("Task2")
    Debug.Print "dict.Exists('Unreal Task'): " & dict.Exists("Unreal Task")

    Dim i As Integer
    For i = LBound(keysArray) To UBound(keysArray)
        Debug.Print keysArray(i) & " -> " & valuesArray(i)
    Next i

    Debug.Print "dict.Count: " & dict.Count

    dict.Remove "Task1"
    
    Debug.Print "dict.Count: " & dict.Count

    dict.RemoveAll
    
    Debug.Print "dict.Count: " & dict.Count
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

Sub TestSimpleDictionary()
    Dim objDict As clsSimpleDictionary
    Set objDict = New clsSimpleDictionary
    
    Debug.Print "=== Initial Status ==="
    Debug.Print "Count: " & objDict.Count
    
    ' Add
    objDict.Add "apple", "red"
    objDict.Add "banana", "yellow"
    objDict.Add "cherry", "pink"
    
    Debug.Print "=== After Add ==="
    Debug.Print "Count: " & objDict.Count
    Debug.Print "apple = " & objDict.GetValue("apple")
    Debug.Print "banana = " & objDict.GetValue("banana")
    
    ' Add to update
    objDict.Add "apple", "strong red"
    Debug.Print "apple(updated) = " & objDict.GetValue("apple")
    
    ' Update with existing key
    objDict.Update "banana", "yellow (updated)"
    Debug.Print "banana(updated) = " & objDict.GetValue("banana")
    
    ' Exists
    If objDict.Exists("cherry") Then
        Debug.Print "cherry exists"
    End If
    
    ' Remove
    objDict.Remove "cherry"
    Debug.Print "cherry removed"
    Debug.Print "Count: " & objDict.Count
    
    ' Clear
    objDict.Clear
    Debug.Print "=== After Clear ==="
    Debug.Print "Count: " & objDict.Count

    ' Error test(GetValue)
    On Error Resume Next
    Dim tempVal As Variant
    tempVal = objDict.GetValue("not_exist")
    If Err.Number <> 0 Then
        Debug.Print "Error(GetValue): " & Err.Description
        Err.Clear
    End If

    ' Error test(Update)
    objDict.Update "not_exist", "xxx"
    If Err.Number <> 0 Then
        Debug.Print "Error(Update): " & Err.Description
        Err.Clear
    End If
    
    Set objDict = Nothing
End Sub


