Attribute VB_Name = "modTestClsWbsNode"
Option Explicit

Sub TestClsWbsNode()
    
    Dim objRoot As New clsWbsNode
    ' objRoot.Id = "T-000"
    objRoot.Name = "Test Project"
    objRoot.StartDate = DateSerial(2025, 4, 1)
    objRoot.EndDate = DateSerial(2025, 4, 30)

    Debug.Print "=== Root Node ==="
    Debug.Print "ID: " & objRoot.Id
    Debug.Print "Name: " & objRoot.Name
    Debug.Print "StartDate: " & objRoot.StartDate
    Debug.Print "EndDate: " & objRoot.EndDate
    Debug.Print "Children Count: " & objRoot.children.Count

    Dim objChild1 As New clsWbsNode
    ' objChild1.Id = "T-001"
    objChild1.Name = "Task 1"
    objChild1.StartDate = DateSerial(2025, 4, 2)
    objChild1.EndDate = DateSerial(2025, 4, 6)

    Dim objChild2 As New clsWbsNode
    ' objChild2.Id = "T-002"
    objChild2.Name = "Task 2"
    objChild2.StartDate = DateSerial(2025, 4, 7)
    objChild2.EndDate = DateSerial(2025, 4, 11)

    objRoot.AddChild objChild1
    objRoot.AddChild objChild2

    Debug.Print vbCrLf & "=== After Adding Children ==="
    Debug.Print "Children Count: " & objRoot.children.Count

    Dim objChild As clsWbsNode
    For Each objChild In objRoot.children
        Debug.Print "  - Child ID: " & objChild.Id
        Debug.Print "    Name: " & objChild.Name
        Debug.Print "    StartDate: " & objChild.StartDate
        Debug.Print "    EndDate: " & objChild.EndDate
        Debug.Print "    Parent Name: " & objChild.Parent.Name
    Next

    Dim objDuplicate As clsWbsNode
    Set objDuplicate = objChild1
    objRoot.AddChild objDuplicate

    Debug.Print vbCrLf & "=== After Attempting Duplicate Child ==="
    Debug.Print "Children Count (should still be 2): " & objRoot.children.Count

    objRoot.AddChild Nothing

    Debug.Print vbCrLf & "=== After Adding Nothing ==="
    Debug.Print "Children Count (should still be 2): " & objRoot.children.Count

    Dim objGrandChild As New clsWbsNode
    ' objGrandChild.Id = "T-003"
    objGrandChild.Name = "Task 3"
    objChild2.AddChild objGrandChild

    Debug.Print vbCrLf & "=== Grandchild Test ==="
    Debug.Print "Grandchild ID: " & objGrandChild.Id
    Debug.Print "Grandchild Name: " & objGrandChild.Name
    Debug.Print "Parent Name: " & objGrandChild.Parent.Name
    Debug.Print "Root's second child (" & objChild2.Name & ") has " & objChild2.children.Count & " child(ren)"

End Sub

