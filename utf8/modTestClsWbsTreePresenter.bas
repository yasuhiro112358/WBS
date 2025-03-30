Attribute VB_Name = "modTestClsWbsTreePresenter"
Option Explicit

Sub TestClsWbsTreePresenter()
    On Error GoTo ErrorHandler

    Dim objTree As New clsWbsTree
    
    Dim objNodeA As New clsWbsNode
    Dim objNodeB As New clsWbsNode
    Dim objNodeC As New clsWbsNode

    objNodeA.Name = "Phase A"
    objNodeB.Name = "Task A-1"
    objNodeC.Name = "Phase B"

    objTree.AddNode objTree.Root.Id, objNodeA
    objTree.AddNode objNodeA.Id, objNodeB
    objTree.AddNode objTree.Root.Id, objNodeC
    
    Dim objPresenter As New clsWbsTreePresenter
    Call objPresenter.InjectTree(objTree)

    Dim objWorksheet As Worksheet
    Set objWorksheet = shtTest2

    Call objPresenter.ExportToSheet(objWorksheet)

    Call LogMessage("Complete exporting to new sheet.")
    
    Exit Sub
    
ErrorHandler:
    Call LogError
    Resume Next
    
End Sub

