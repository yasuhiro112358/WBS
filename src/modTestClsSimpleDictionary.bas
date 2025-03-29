Attribute VB_Name = "modTestClsSimpleDictionary"
Option Explicit

Sub TestClsSimpleDictionary()
    On Error GoTo ErrorHandler

    Dim objDict As clsSimpleDictionary
    Set objDict = New clsSimpleDictionary
    
    Debug.Print "=== Initial Status ==="
    Debug.Print "Count: " & objDict.Count
    
    
    objDict.Add "apple", "red"
    objDict.Add "banana", "yellow"
    objDict.Add "cherry", "pink"
    
    Debug.Print "=== After adding string values ==="
    Debug.Print "Count: " & objDict.Count
    Debug.Print "apple = " & objDict.GetValue("apple")
    Debug.Print "banana = " & objDict.GetValue("banana")
    Debug.Print "cherry = " & objDict.GetValue("cherry")
    
    Dim objNode As clsWbsNode
    Set objNode = New clsWbsNode
    ' Debug.Print "TypeName:", TypeName(objNode)
    ' Debug.Print "objNode.Id:", objNode.Id
    objDict.Add "Test Object", objNode
    
    Debug.Print "=== After adding an object ==="
    Dim Check As Boolean
    Check = objDict.Exists("Test Object") ' True
    Debug.Print "objDict.Exists('Test Object'): " & Check
    
    Dim AddedObject As clsWbsNode
    Set AddedObject = objDict.GetValue("Test Object")
    Debug.Print "TypeName(AddedObject): " & TypeName(AddedObject)
    
    ' Get key or value by index
    Debug.Print "Key of Index(0): " & objDict.GetKeyByIndex(0)
    Debug.Print "Value of Index(1): " & objDict.GetValueByIndex(1)
    
    ' Make errors
    Debug.Print "Trying 'Key of Index(-1)'"
    Debug.Print "Key of Index(-1): " & objDict.GetKeyByIndex(-1)
    
    Debug.Print "Trying 'Key of Index(-2)'"
    Debug.Print "Value of Index(-2): " & objDict.GetValueByIndex(-2)
      
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
    
    Exit Sub
    
ErrorHandler:
    Call LogError
    Resume Next

End Sub
