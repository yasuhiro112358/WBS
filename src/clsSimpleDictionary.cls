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
    Erase m_arrEntries
    m_EntryCount = 0
End Sub

Public Sub Add(ByVal p_DictKey As String, ByVal p_DictValue As Variant)
    Dim Index As Long
    Index = GetIndexByKey(p_DictKey)
    If Index <> -1 Then
        If IsObject(p_DictValue) Then
            Set m_arrEntries(Index).DictValue = p_DictValue
        Else
            m_arrEntries(Index).DictValue = p_DictValue
        End If
    Else
        ReDim Preserve m_arrEntries(0 To m_EntryCount)
        
        m_arrEntries(m_EntryCount).DictKey = p_DictKey
        If IsObject(p_DictValue) Then
            Set m_arrEntries(m_EntryCount).DictValue = p_DictValue
        Else
            m_arrEntries(m_EntryCount).DictValue = p_DictValue
        End If
        
        m_EntryCount = m_EntryCount + 1
    End If
End Sub

Public Function Exists(ByVal p_DictKey As String) As Boolean
    Exists = (GetIndexByKey(p_DictKey) <> -1)
End Function

Public Function GetValue(ByVal p_DictKey As String) As Variant
    Dim Index As Long
    Index = GetIndexByKey(p_DictKey)
    If Index = -1 Then
        RaiseError DictError.KeyNotFound, "GetValue", "Key not found: " & p_DictKey
        Exit Function
    End If

    If IsObject(m_arrEntries(Index).DictValue) Then
        Set GetValue = m_arrEntries(Index).DictValue
    Else
        GetValue = m_arrEntries(Index).DictValue
    End If
End Function

Public Sub Update(ByVal p_DictKey As String, ByVal p_DictValue As Variant)
    Dim Index As Long
    Index = GetIndexByKey(p_DictKey)
    If Index = -1 Then
        RaiseError DictError.KeyNotFound, "GetValue", "Key not found: " & p_DictKey
        Exit Sub
    End If
    
    If IsObject(p_DictValue) Then
        Set m_arrEntries(Index).DictValue = p_DictValue
    Else
        m_arrEntries(Index).DictValue = p_DictValue
    End If
End Sub

Public Sub Remove(ByVal p_DictKey As String)
    Dim Index As Long
    Index = GetIndexByKey(p_DictKey)
    If Index = -1 Then
        Err.Raise vbObjectError + 1004, , "Key not found: " & p_DictKey
        Exit Sub
    End If
    
    Dim i As Long
    For i = Index To m_EntryCount - 2
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
Private Function GetIndexByKey(ByVal p_DictKey As String) As Long
    Dim i As Long
    For i = 0 To m_EntryCount - 1
        If m_arrEntries(i).DictKey = p_DictKey Then
            GetIndexByKey = i
            Exit Function
        End If
    Next i
    GetIndexByKey = -1
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
        Call RaiseError(DictError.IndexOutOfRange, p_Source, "Index out of range")
    End If
End Sub
