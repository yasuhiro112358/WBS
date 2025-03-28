VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWbsNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Id As String
Private m_Name As String
Private m_StartDate As Date
Private m_EndDate As Date
Private m_objParent As clsWbsNode
Private m_arrChildren As Collection

Private Sub Class_Initialize()
    m_Id = GenerateUUIDv4()
    m_Name = ""
    m_StartDate = 0
    m_EndDate = 0
    Set m_objParent = Nothing
    Set m_arrChildren = New Collection
End Sub

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(p_Name As String)
    m_Name = p_Name
End Property

Public Property Get StartDate() As Date
    StartDate = m_StartDate
End Property

Public Property Let StartDate(p_StartDate As Date)
    m_StartDate = p_StartDate
End Property

Public Property Get EndDate() As Date
    EndDate = m_EndDate
End Property

Public Property Let EndDate(p_EndDate As Date)
    m_EndDate = p_EndDate
End Property

Public Property Get Parent() As clsWbsNode
    Set Parent = m_objParent
End Property

Public Property Get Children() As Collection
   Set Children = m_arrChildren
End Property

'
' Only for unit test
'
Public Sub AddChild(p_objChild As clsWbsNode)
    If p_objChild Is Nothing Then
        Exit Sub
    End If

    Dim objChild As clsWbsNode
    For Each objChild In m_arrChildren
        If objChild Is p_objChild Then
            Exit Sub
        End If
    Next objChild

    m_arrChildren.Add p_objChild
    Set p_objChild.Parent = Me
End Sub

'
' Only for unit test
'
Public Property Set Parent(p_objParent As clsWbsNode)
    Set m_objParent = p_objParent
End Property

