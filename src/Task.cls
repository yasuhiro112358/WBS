VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Task"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private id_ As String
Private name_ As String
Private start_date_ As Date
Private end_date_ As Date
Private parent_ As Task
Private children_ As Collection

Private Sub Class_Initialize()
    id_ = ""
    name_ = ""
    start_date_ = 0
    end_date_ = 0
    Set children_ = New Collection
End Sub

Public Sub Initialize(ByVal name__ As String, ByVal start_date__ As Date, ByVal end_date__ As Date)
    id_ = GenerateUUIDv4()
    name_ = name__
    start_date_ = start_date__
    end_date_ = end_date__
End Sub

Public Property Get Id() As String
    Id = id_
End Property

Public Property Get Name() As String
    Name = name_
End Property

Public Property Get StartDate() As Date
    StartDate = start_date_
End Property

Public Property Get EndDate() As Date
    EndDate = end_date_
End Property

Public Property Get Parent() As Task
    Set Parent = parent_
End Property

Public Property Set Parent(ByVal arg As Task)
    Set Parent = arg
End Property

Public Sub AddChild(ByVal child As Task)
    children_.Add child
    Set child.Parent = Me
End Sub

Public Function GetChildren() As Collection
    Set GetChildren = children_
End Function

Public Sub WriteToSheet(ByVal ws As Worksheet, ByVal row As Integer)
    ws.Cells(row, 1).value = id_
    ws.Cells(row, 2).value = name_
    ws.Cells(row, 3).value = start_date_
    ws.Cells(row, 4).value = end_date_
End Sub

