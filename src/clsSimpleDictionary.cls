VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSimpleDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type DictEntry
    DictKey As String
    DictValue As Variant
End Type

Private m_arrEntries() As DictEntry
Private m_EntryCount As Long

Private Sub Class_Initialize()
    ' ReDim m_arrEntries(0 To -1)
    Erase m_arrEntries
    m_EntryCount = 0
End Sub

Public Sub Add(ByVal p_DictKey As String, ByVal p_DictValue As Variant)
    Dim KeyIndex As Long
    KeyIndex = GetKeyIndex(p_DictKey)

    If KeyIndex <> -1 Then
        m_arrEntries(KeyIndex).DictValue = p_DictValue
    Else
        ReDim Preserve m_arrEntries(0 To m_EntryCount)
        m_arrEntries(m_EntryCount).DictKey = p_DictKey
        m_arrEntries(m_EntryCount).DictValue = p_DictValue
        m_EntryCount = m_EntryCount + 1
    End If
End Sub

Public Function Exists(ByVal p_DictKey As String) As Boolean
    Exists = (GetKeyIndex(p_DictKey) <> -1)
End Function

Public Function GetValue(ByVal p_DictKey As String) As Variant
    Dim KeyIndex As Long
    KeyIndex = GetKeyIndex(p_DictKey)

    If KeyIndex = -1 Then
        Err.Raise vbObjectError + 1003, , "Key not found: " & p_DictKey
    End If

    GetValue = m_arrEntries(KeyIndex).DictValue
End Function

Public Sub Update(ByVal p_DictKey As String, ByVal p_DictValue As Variant)
    Dim KeyIndex As Long
    KeyIndex = GetKeyIndex(p_DictKey)

    If KeyIndex = -1 Then
        Err.Raise vbObjectError + 1005, , "Key not found: " & p_DictKey
    End If

    m_arrEntries(KeyIndex).DictValue = p_DictValue
End Sub

Public Sub Remove(ByVal p_DictKey As String)
    Dim KeyIndex As Long
    Dim i As Long
    
    KeyIndex = GetKeyIndex(p_DictKey)

    If KeyIndex = -1 Then
        Err.Raise vbObjectError + 1004, , "Key not found: " & p_DictKey
    End If

    For i = KeyIndex To m_EntryCount - 2
        m_arrEntries(i) = m_arrEntries(i + 1)
    Next i

    m_EntryCount = m_EntryCount - 1
    If m_EntryCount = 0 Then
        ReDim m_arrEntries(0 To -1)
    Else
        ReDim Preserve m_arrEntries(0 To m_EntryCount - 1)
    End If
End Sub

Public Property Get Count() As Long
    Count = m_EntryCount
End Property

Public Sub Clear()
    Erase m_arrEntries
    m_EntryCount = 0
End Sub

'
' return: Index number, if not found return `-1`
'
Private Function GetKeyIndex(ByVal p_DictKey As String) As Long
    Dim i As Long
    For i = 0 To m_EntryCount - 1
        If m_arrEntries(i).DictKey = p_DictKey Then
            GetKeyIndex = i
            Exit Function
        End If
    Next i
    GetKeyIndex = -1
End Function

Public Function GetKeyByIndex(ByVal p_Index As Long) As String
    Call ValidateIndex(p_Index, "GetKeyByIndex")
    GetKeyByIndex = m_arrEntries(p_Index).DictKey
End Function

Public Function GetValueByIndex(ByVal p_Index As Long) As Variant
    Call ValidateIndex(p_Index, "GetValueByIndex")
    GetValueByIndex = m_arrEntries(p_Index).DictValue
End Function

Private Sub ValidateIndex(ByVal p_Index As Long, ByVal p_Source As String)
    If p_Index < 0 Or p_Index >= m_EntryCount Then
'        Err.Raise vbObjectError + p_ErrorCode, , _
'            "Index out of range in " & p_Context & ": " & p_Index
        Call RaiseError(DictError.IndexOutOfRange, p_Source, "Index out of range")
    End If
End Sub
