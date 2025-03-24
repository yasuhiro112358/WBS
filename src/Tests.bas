Attribute VB_Name = "Tests"
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
    Dim task1 As Task
    Set task1 = New Task
    task1.Initialize "Hiro", Date, "2025-4-1"
    
    Debug.Print task1.Id
    Debug.Print task1.Name
    Debug.Print task1.StartDate
    Debug.Print task1.EndDate
End Sub
