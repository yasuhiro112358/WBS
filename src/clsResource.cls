VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Id As String
Private m_Name As String

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(p_Value As String)
    m_Id = p_Value
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(p_Value As String)
    m_Name = p_Value
End Property

