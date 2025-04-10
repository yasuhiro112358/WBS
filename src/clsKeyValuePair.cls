VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKeyValuePair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Key As String
Private m_Value As Variant

Public Property Get Key() As String
    Key = m_Key
End Property

Public Property Let Key(ByVal p_Key As String)
    m_Key = p_Key
End Property

Public Property Get Value() As Variant
    If IsObject(m_Value) Then
        Set Value = m_Value
    Else
        Value = m_Value
    End If
End Property

Public Property Let Value(ByVal p_Value As Variant)
    If IsObject(p_Value) Then
        Err.Raise 13, , "Use Set for object values."
    End If
    m_Value = p_Value
End Property

Public Property Set Value(ByVal p_Value As Variant)
    Set m_Value = p_Value
End Property

