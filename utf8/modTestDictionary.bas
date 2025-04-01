Attribute VB_Name = "modTestDictionary"
Option Explicit

Public Sub TestDictionary()
    Dim Dict As clsDictionary
    Set Dict = New clsDictionary
    
    Debug.Print "== Add test =="
    Dict.Add "A", 100
    Dict.Add "B", 200
    Dict.Add "C", "Text"
    Dict.Add "D", Date
    Dict.Add "E", True
    Debug.Assert Dict.Count = 5
    Debug.Print "Add test passed"
    
    Debug.Print "== Add Object test =="
    Dim objWorksheet As Worksheet
    Set objWorksheet = shtTest
    Dict.Add "Sheet1", objWorksheet
    Debug.Assert Dict.Item("Sheet1").Name = objWorksheet.Name
    Debug.Print "Object test passed: shtTest name = " & Dict.Item("Sheet1").Name
    
    Debug.Print "== Duplicate Add test =="
    On Error Resume Next
    Dict.Add "A", 300
    Debug.Assert Err.Number <> 0
    Err.Clear
    On Error GoTo 0
    Debug.Print "Duplicate Add raised expected error"
    
    Debug.Print "== Item test =="
    Debug.Assert Dict.Item("A") = 100
    Debug.Assert Dict.Item("B") = 200
    Debug.Print "Item test passed"
    
    Debug.Print "== Update test =="
    Dict.Update "A", 150
    Debug.Assert Dict.Item("A") = 150
    On Error Resume Next
    Dict.Update "X", 999
    Debug.Assert Err.Number <> 0
    Err.Clear
    On Error GoTo 0
    Debug.Print "Update test passed"
    
    Debug.Print "== Exists test =="
    Debug.Assert Dict.Exists("A") = True
    Debug.Assert Dict.Exists("Z") = False
    Debug.Print "Exists test passed"
    
    Debug.Print "== Remove test =="
    Dict.Remove "A"
    Debug.Assert Dict.Exists("A") = False
    On Error Resume Next
    Dict.Remove "A"
    Debug.Assert Err.Number <> 0
    Err.Clear
    On Error GoTo 0
    Debug.Print "Remove test passed"
    
    Debug.Print "== Keys / Values test =="
    Dim Key As Variant
    Dim Val As Variant
    For Each Key In Dict.Keys
        Debug.Print "Key: " & Key
    Next
    For Each Val In Dict.Values
        If IsObject(Val) Then
            Debug.Print "Value (Object): " & TypeName(Val)
        Else
            Debug.Print "Value: " & Val
        End If
    Next
    Debug.Print "Keys / Values test passed"
    
    Debug.Print "== Clear test =="
    Dict.Clear
    Debug.Assert Dict.Count = 0
    Debug.Print "Clear test passed"
    
    Debug.Print "== All tests completed =="
End Sub

