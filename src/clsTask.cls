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

Private mId As String
Private mName As String
Private mStartDate As Date
Private mEndDate As Date
Private mParent As clsTask
Private mChildren As Collection

Private Sub Class_Initialize()
    mId = ""
    mName = ""
    mStartDate = 0
    mEndDate = 0
    Set mChildren = New Collection
End Sub

Public Sub Initialize(ByVal pName As String, ByVal pStartDate As Date, ByVal pEndDate As Date)
    mId = GenerateUUIDv4()
    mName = pName
    mStartDate = pStartDate
    mEndDate = pEndDate
End Sub

Public Property Get Id() As String
    Id = mId
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal pName As String)
    mName = pName
End Property

Public Property Get StartDate() As Date
    StartDate = mStartDate
End Property

Public Property Get EndDate() As Date
    EndDate = mEndDate
End Property

Public Property Get Parent() As clsTask
    Set Parent = mParent
End Property

Public Property Set Parent(ByVal pParent As clsTask)
    Set mParent = pParent
End Property

Public Sub AddChild(ByVal pChild As clsTask)
    mChildren.Add pChild
    Set pChild.Parent = Me
End Sub

Public Function GetChildren() As Collection
    Set GetChildren = mChildren
End Function

Public Sub WriteToSheet(ByVal ws As Worksheet, ByVal row As Integer)
    ws.Cells(row, TASK_ID).value = mId
    ws.Cells(row, TASK_NAME).value = mName
    ws.Cells(row, TASK_START_DATE).value = mStartDate
    ws.Cells(row, TASK_END_DATE).value = mEndDate
End Sub

