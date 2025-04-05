Attribute VB_Name = "modTestClsWbsTreeStorage"
Option Explicit

Sub TestClsWbsTreeStorage()
    On Error GoTo ErrorHandler

    Dim objRoot As New clsWbsNode
    objRoot.Id = GenerateUUIDv4()
    objRoot.Name = "Test Project"
    
    Dim objNodeA As New clsWbsNode
    objNodeA.Id = GenerateUUIDv4()
    objNodeA.Name = "Phase A"
    
    Dim objNodeB As New clsWbsNode
    objNodeB.Id = GenerateUUIDv4()
    objNodeB.Name = "Task A-1"
    
    Dim objNodeC As New clsWbsNode
    objNodeC.Id = GenerateUUIDv4()
    objNodeC.Name = "Phase B"
    
    Dim objTree As New clsWbsTree
    
    Set objTree.Root = objRoot
    objTree.AddNode "", objRoot
    
    objTree.AddNode objTree.Root.Id, objNodeA
    objTree.AddNode objNodeA.Id, objNodeB
    objTree.AddNode objTree.Root.Id, objNodeC

    Dim objWorksheet As Worksheet
    Set objWorksheet = shtDb
    
    Dim objStorage As New clsWbsTreeStorage
    objStorage.SaveToSheet objWorksheet, objTree

    Call LogMessage("Complete saving tree data")
    
    Exit Sub
    
ErrorHandler:
    Call HandleError("modTestClsWbsTreeStorage", "TestClsWbsTreeStorage")
    
End Sub

Sub TestClsWbsTreeStorage2()
    On Error GoTo ErrorHandler
    
    Dim objStorage As New clsWbsTreeStorage
    
    Dim objWorksheet As Worksheet
    Set objWorksheet = shtDb
    
    Dim objTree As clsWbsTree
    Set objTree = objStorage.LoadFromSheet(objWorksheet)
    
    Call LogMessage("Complete loading.")

    Dim objNewNode As New clsWbsNode
    objNewNode.Id = GenerateUUIDv4()
    objNewNode.Name = "New Task from VBA"
    objTree.AddNode objTree.Root.Id, objNewNode
    
    Call LogMessage("Added new node.")

    Call objStorage.SaveToSheet(objWorksheet, objTree)
    Call LogMessage("Complete saving.")
    
    Exit Sub
    
ErrorHandler:
    Call HandleError("modTestClsWbsTreeStorage", "TestClsWbsTreeStorage2")

End Sub

