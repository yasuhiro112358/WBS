VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWbsTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_objRoot As clsWbsNode
Private m_dictNodes As clsSimpleDictionary
Private m_dictParent As clsSimpleDictionary

Private Sub Class_Initialize()
    Set m_objRoot = New clsWbsNode
    Set m_dictNodes = New clsSimpleDictionary
    Set m_dictParent = New clsSimpleDictionary
    
    m_dictNodes.Add m_objRoot.Id, m_objRoot
End Sub

Public Property Get Root() As clsWbsNode
    Set Root = m_objRoot
End Property

Public Property Set Root(p_objRoot As clsWbsNode)
    Set m_objRoot = p_objRoot
End Property

Public Sub AddNode(p_ParentId As String, p_objChild As clsWbsNode)
    If p_objChild Is Nothing Then
        Exit Sub
    End If
    
    If Not m_dictNodes.Exists(p_objChild.Id) Then
        m_dictNodes.Add p_objChild.Id, p_objChild
    End If

    If Not m_dictParent.Exists(p_objChild.Id) Then
        m_dictParent.Add p_objChild.Id, p_ParentId
    End If
End Sub

Public Function GetParent(p_NodeId As String) As clsWbsNode
    If m_dictParent.Exists(p_NodeId) Then
        Dim ParentId As String
        ParentId = m_dictParent.GetValue(p_NodeId)
        Set GetParent = m_dictNodes.GetValue(ParentId)
    Else
        Set GetParent = Nothing
    End If
End Function

'
' return: Collection[Objects]
'
Public Function GetChildren(p_ParentId As String) As Collection
    Dim colChildren As Collection
    Set colChildren = New Collection
    
    Dim i As Long
    For i = 0 To m_dictParent.Count - 1
        If m_dictParent.GetValueByIndex(i) = p_ParentId Then
            Dim ChildId As String
            ChildId = m_dictParent.GetKeyByIndex(i)
            
            colChildren.Add m_dictNodes.GetValue(ChildId)
        End If
    Next i

    Set GetChildren = colChildren
End Function

Public Function FindNodeById(p_Id As String) As clsWbsNode
    If m_dictNodes.Exists(p_Id) Then
        Set FindNodeById = m_dictNodes.GetValue(p_Id)
    Else
        Set FindNodeById = Nothing
    End If
End Function

