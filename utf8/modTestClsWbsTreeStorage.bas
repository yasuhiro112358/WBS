Attribute VB_Name = "modTestClsWbsTreeStorage"
Option Explicit

Sub TestClsWbsTreeStorage()
    Dim objTree As New clsWbsTree
    Dim objNodeA As New clsWbsNode
    Dim objNodeB As New clsWbsNode
    Dim objNodeC As New clsWbsNode

    objTree.Root.Name = "Project"
    
    objNodeA.Name = "Phase A"
    objNodeB.Name = "Task A-1"
    objNodeC.Name = "Phase B"

    objTree.AddNode objTree.Root.Id, objNodeA
    objTree.AddNode objNodeA.Id, objNodeB
    objTree.AddNode objTree.Root.Id, objNodeC

    Dim objStorage As New clsWbsTreeStorage

    Dim objWorksheet As Worksheet
    Set objWorksheet = shtDb
    
    Call objStorage.SaveToSheet(objWorksheet, objTree)

    Call LogMessage("Complete saving tree data")
End Sub

Sub TestClsWbsTreeStorage2()
    On Error GoTo ErrorHandler

    Dim objWorksheet As Worksheet
    Set objWorksheet = shtDb

    Dim objStorage As New clsWbsTreeStorage
    Dim objTree As clsWbsTree
    Set objTree = objStorage.LoadFromSheet(objWorksheet)
    Call LogMessage("Complete loading.")

    Dim objNewNode As New clsWbsNode
    objNewNode.Name = "New Task from VBA"
    objTree.AddNode objTree.Root.Id, objNewNode
    Call LogMessage("Added new node.")

    Call objStorage.SaveToSheet(objWorksheet, objTree)
    Call LogMessage("Complete saving.")
    
    Exit Sub
    
ErrorHandler:
    Call LogError
    Resume Next

End Sub

