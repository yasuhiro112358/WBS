Attribute VB_Name = "modTests"
Option Explicit

Sub TestGenerateUUIDv4()
   Dim uuid As String
   uuid = GenerateUUIDv4()
   Debug.Print uuid
End Sub

Sub TestCustomDictionary()
    Debug.Print "Start TestCustomDictionary"
    
    Dim dict As New CustomDictionary
    
    Dim keysArray As Variant
    Dim valuesArray As Variant
        
    ' Add Tasks
    dict.Add "Task1", "Write VBA Code"
    dict.Add "Task2", "Review Code"
    dict.Item("Task3") = "Test Code"
    
    ' Get all keys
    keysArray = dict.Keys
    Dim keyArray As Variant
    For Each keyArray In keysArray
        Debug.Print "keyArray: " & keyArray
    Next
    
    ' Get all values
    valuesArray = dict.Items
    Dim valueArray As Variant
    For Each valueArray In valuesArray
        Debug.Print "valueArray: " & valueArray
    Next

    Debug.Print "Check if a key exist or not"
    Debug.Print "dict.Exists('Task2'): " & dict.Exists("Task2")
    Debug.Print "dict.Exists('Unreal Task'): " & dict.Exists("Unreal Task")

    ' ???????????
    Dim key As Variant
    'Dim keysArray As Variant
    
    
    
    keysArray = dict.Keys()
    valuesArray = dict.Items()

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

