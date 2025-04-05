VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWbsManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_objTree As clsWbsTree

Private Sub Class_Initialize()
    Set m_objTree = New clsWbsTree
End Sub

Public Sub CreateProject(p_Name As String)
    Dim objRootNode As clsWbsNode
    Set objRootNode = New clsWbsNode
    objRootNode.Id = GenerateUUIDv4()
    objRootNode.Name = p_Name
    Set objRootNode.Parent = Nothing
    
    m_objTree.Root = objRootNode
End Sub

Public Sub CreateNode(p_ParentNodeId As String, p_NodeName As String)
    ' If parent does not exit going to error

    Dim objNewNode As clsWbsNode
    Set objNewNode = New clsWbsNode
    objNewNode.Id = GenerateUUIDv4()
    objNewNode.Name = p_NodeName
    
    Dim objParentNode As clsWbsNode
    Set objParentNode = m_objTree.FindNodeById(p_ParentNodeId)
    objParentNode.AddChild objNewNode
End Sub

Public Sub LoadFromSheet(ws As Worksheet)
    ' Excel????WbsTree???
End Sub

Public Sub WriteToSheet(ws As Worksheet)
    ' WbsTree????Excel???
End Sub
