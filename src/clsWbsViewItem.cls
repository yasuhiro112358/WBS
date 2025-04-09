VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWbsViewItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_TaskId As String
Private m_WbsCode As String
Private m_Name As String
Private m_Level As Long
Private m_DisplayName As String
Private m_StartDate As Date
Private m_EndDate As Date
Private m_Progress As Double
Private m_Duration As Double
Private m_BaselineStartDate As Date
Private m_BaselineEndDate As Date
Private m_BaselineWorkHours As Double
Private m_AssignedWorkHours As Double
Private m_ActualWorkHours As Double
Private m_RemainingWorkHours As Double
Private m_ActualStartDate As Date
Private m_ActualEndDate As Date
Private m_PredecessorId As String
Private m_PredecessorWbsCode As String
Private m_PredecessorName As String
Private m_ResourceId As String
Private m_ResourceName As String
Private m_StatusLabel As String
Private m_TaskLineNo As Long
Private m_ParentLineNo As Long
Private m_PredecessorLineNo As Long

Public Property Get TaskId() As String
    TaskId = m_TaskId
End Property

Public Property Let TaskId(p_Value As String)
    m_TaskId = p_Value
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

Public Property Get Level() As Long
    Level = m_Level
End Property

Public Property Let Level(p_Value As Long)
    m_Level = p_Value
End Property

Public Property Get DisplayName() As String
    DisplayName = m_DisplayName
End Property

Public Property Let DisplayName(p_Value As String)
    m_DisplayName = p_Value
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

Public Property Get Progress() As Double
    Progress = m_Progress
End Property

Public Property Let Progress(p_Value As Double)
    m_Progress = p_Value
End Property

Public Property Get Duration() As Double
    Duration = m_Duration
End Property

Public Property Let Duration(p_Value As Double)
    m_Duration = p_Value
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

Public Property Get AssignedWorkHours() As Double
    AssignedWorkHours = m_AssignedWorkHours
End Property

Public Property Let AssignedWorkHours(p_Value As Double)
    m_AssignedWorkHours = p_Value
End Property

Public Property Get ActualWorkHours() As Double
    ActualWorkHours = m_ActualWorkHours
End Property

Public Property Let ActualWorkHours(p_Value As Double)
    m_ActualWorkHours = p_Value
End Property

Public Property Get RemainingWorkHours() As Double
    RemainingWorkHours = m_RemainingWorkHours
End Property

Public Property Let RemainingWorkHours(p_Value As Double)
    m_RemainingWorkHours = p_Value
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

Public Property Get PredecessorId() As String
    PredecessorId = m_PredecessorId
End Property

Public Property Let PredecessorId(p_Value As String)
    m_PredecessorId = p_Value
End Property

Public Property Get PredecessorWbsCode() As String
    PredecessorWbsCode = m_PredecessorWbsCode
End Property

Public Property Let PredecessorWbsCode(p_Value As String)
    m_PredecessorWbsCode = p_Value
End Property

Public Property Get PredecessorName() As String
    PredecessorName = m_PredecessorName
End Property

Public Property Let PredecessorName(p_Value As String)
    m_PredecessorName = p_Value
End Property

Public Property Get ResourceId() As String
    ResourceId = m_ResourceId
End Property

Public Property Let ResourceId(p_Value As String)
    m_ResourceId = p_Value
End Property

Public Property Get ResourceName() As String
    ResourceName = m_ResourceName
End Property

Public Property Let ResourceName(p_Value As String)
    m_ResourceName = p_Value
End Property

Public Property Get StatusLabel() As String
    StatusLabel = m_StatusLabel
End Property

Public Property Let StatusLabel(p_Value As String)
    m_StatusLabel = p_Value
End Property

Public Property Get TaskLineNo() As Long
    TaskLineNo = m_TaskLineNo
End Property

Public Property Let TaskLineNo(p_Value As Long)
    m_TaskLineNo = p_Value
End Property

Public Property Get ParentLineNo() As Long
    ParentLineNo = m_ParentLineNo
End Property

Public Property Let ParentLineNo(p_Value As Long)
    m_ParentLineNo = p_Value
End Property

Public Property Get PredecessorLineNo() As Long
    PredecessorLineNo = m_PredecessorLineNo
End Property

Public Property Let PredecessorLineNo(p_Value As Long)
    m_PredecessorLineNo = p_Value
End Property

