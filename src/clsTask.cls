VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Id As String
Private m_Name As String
Private m_StartDate As Date
Private m_EndDate As Date
Private m_Parent As clsTask
Private m_Children As Collection

Private Sub Class_Initialize()
    m_Id = ""
    m_Name = ""
    m_StartDate = 0
    m_EndDate = 0
    Set m_Children = New Collection
End Sub

Public Sub Initialize(ByVal p_Name As String, ByVal p_StartDate As Date, ByVal p_EndDate As Date)
    m_Id = GenerateUUIDv4()
    m_Name = p_Name
    m_StartDate = p_StartDate
    m_EndDate = p_EndDate
End Sub

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(ByVal p_Name As String)
    m_Name = p_Name
End Property

Public Property Get StartDate() As Date
    StartDate = m_StartDate
End Property

Public Property Get EndDate() As Date
    EndDate = m_EndDate
End Property

Public Property Get Parent() As clsTask
    Set Parent = m_Parent
End Property

Public Property Set Parent(ByVal p_Parent As clsTask)
    Set m_Parent = p_Parent
End Property

Public Sub AddChild(ByVal p_Child As clsTask)
    m_Children.Add p_Child
    Set p_Child.Parent = Me
End Sub

Public Function GetChildren() As Collection
    Set GetChildren = m_Children
End Function

Public Sub WriteToSheet(ByVal ws As Worksheet, ByVal row As Integer)
    ws.Cells(row, TASK_ID).value = m_Id
    ws.Cells(row, TASK_NAME).value = m_Name
    ws.Cells(row, TASK_START_DATE).value = m_StartDate
    ws.Cells(row, TASK_END_DATE).value = m_EndDate
End Sub

