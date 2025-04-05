Attribute VB_Name = "modWbsTests"
Option Explicit

Sub TestWbsManager()
'    Dim objRootNode As clsWbsNode
'    Set objRootNode = New clsWbsNode
'    objRootNode.Id = GenerateUUIDv4()
'    objRootNode.Name = "Test Project (this is the root node)"
'    Set objRootNode.Parent = Nothing
'
'
'
'    Debug.Print objRootNode.Id
'    Debug.Print objRootNode.Name
'
'    Dim objTree As clsWbsTree
'    Set objTree = New clsWbsTree
'    Set objTree.Root = objRootNode
    
    

    Dim objWbsManager As clsWbsManager
    Set objWbsManager = New clsWbsManager
    objWbsManager.CreateProject "Test Project (this is the root node)"
    
    
    ' objWbsManager.CreateNode objRootNode.Id, "Test Task 1"
    
    
    
    
End Sub

