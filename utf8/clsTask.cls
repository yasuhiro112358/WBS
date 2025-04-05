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
Private m_WbsCode As String
Private m_Name As String
Private m_StartDate As Date
Private m_EndDate As Date
Private m_Duration As Double
Private m_AssignedWorkHours As Double
Private m_Progress As Double
Private m_ActualStartDate As Date
Private m_ActualEndDate As Date
Private m_ActualWorkHours As Double
Private m_BaselineStartDate As Date
Private m_BaselineEndDate As Date
Private m_BaselineWorkHours As Double
Private m_PredecessorId As String
Private m_ParentId As String
Private m_ResourceId As String

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(p_Value As String)
    m_Id = p_Value
End Property

Public Property Get WbsCode() As String
    WbsCode = m_WbsCode
End Property

Public Property Let WbsCode(p_Value As String)
    m_WbsCode = p_Value
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(p_Value As String)
    m_Name = p_Value
End Property

Public Property Get StartDate() As Date
    StartDate = m_StartDate
End Property

Public Property Let StartDate(p_Value As Date)
    m_StartDate = p_Value
End Property

Public Property Get EndDate() As Date
    EndDate = m_EndDate
End Property

Public Property Let EndDate(p_Value As Date)
    m_EndDate = p_Value
End Property

Public Property Get Duration() As Double
    Duration = m_Duration
End Property

Public Property Let Duration(p_Value As Double)
    m_Duration = p_Value
End Property

Public Property Get AssignedWorkHours() As Double
    AssignedWorkHours = m_AssignedWorkHours
End Property

Public Property Let AssignedWorkHours(p_Value As Double)
    m_AssignedWorkHours = p_Value
End Property

Public Property Get Progress() As Double
    Progress = m_Progress
End Property

Public Property Let Progress(p_Value As Double)
    m_Progress = p_Value
End Property

Public Property Get ActualStartDate() As Date
    ActualStartDate = m_ActualStartDate
End Property

Public Property Let ActualStartDate(p_Value As Date)
    m_ActualStartDate = p_Value
End Property

Public Property Get ActualEndDate() As Date
    ActualEndDate = m_ActualEndDate
End Property

Public Property Let ActualEndDate(p_Value As Date)
    m_ActualEndDate = p_Value
End Property

Public Property Get ActualWorkHours() As Double
    ActualWorkHours = m_ActualWorkHours
End Property

Public Property Let ActualWorkHours(p_Value As Double)
    m_ActualWorkHours = p_Value
End Property

Public Property Get BaselineStartDate() As Date
    BaselineStartDate = m_BaselineStartDate
End Property

Public Property Let BaselineStartDate(p_Value As Date)
    m_BaselineStartDate = p_Value
End Property

Public Property Get BaselineEndDate() As Date
    BaselineEndDate = m_BaselineEndDate
End Property

Public Property Let BaselineEndDate(p_Value As Date)
    m_BaselineEndDate = p_Value
End Property

Public Property Get BaselineWorkHours() As Double
    BaselineWorkHours = m_BaselineWorkHours
End Property

Public Property Let BaselineWorkHours(p_Value As Double)
    m_BaselineWorkHours = p_Value
End Property

Public Property Get PredecessorId() As String
    PredecessorId = m_PredecessorId
End Property

Public Property Let PredecessorId(p_Value As String)
    m_PredecessorId = p_Value
End Property

Public Property Get ParentId() As String
    ParentId = m_ParentId
End Property

Public Property Let ParentId(p_Value As String)
    m_ParentId = p_Value
End Property

Public Property Get ResourceId() As String
    ResourceId = m_ResourceId
End Property

Public Property Let ResourceId(p_Value As String)
    m_ResourceId = p_Value
End Property

