VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME As String = "clsDictionary"

Private m_colItems As Collection

Private Sub Class_Initialize()
    Set m_colItems = New Collection
End Sub

Public Sub Add(ByVal p_Key As String, ByVal p_Value As Variant)
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "Add"

    If Exists(p_Key) Then
        Err.Raise vbObjectError + 1, MODULE_NAME & "." & PROCEDURE_NAME, "Key already exists: " & p_Key
    End If

    Dim objKeyValuePair As clsKeyValuePair
    Set objKeyValuePair = New clsKeyValuePair
    objKeyValuePair.Key = p_Key
     
    If IsObject(p_Value) Then
        Set objKeyValuePair.Value = p_Value
    Else
        objKeyValuePair.Value = p_Value
    End If
        
    m_colItems.Add objKeyValuePair

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Sub

Public Function Item(ByVal p_Key As String) As Variant
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "Item"

    Dim objKeyValuePair As clsKeyValuePair
    For Each objKeyValuePair In m_colItems
        If objKeyValuePair.Key = p_Key Then
            If IsObject(objKeyValuePair.Value) Then
               Set Item = objKeyValuePair.Value
            Else
                Item = objKeyValuePair.Value
            End If
            Exit Function
        End If
    Next

    Err.Raise vbObjectError + 2, MODULE_NAME & "." & PROCEDURE_NAME, "Key not found: " & p_Key
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

Public Sub Update(ByVal p_Key As String, ByVal p_Value As Variant)
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "Update"

    Dim objKeyValuePair As clsKeyValuePair
    For Each objKeyValuePair In m_colItems
        If objKeyValuePair.Key = p_Key Then
            If IsObject(p_Value) Then
                Set objKeyValuePair.Value = p_Value
            Else
                objKeyValuePair.Value = p_Value
            End If
            Exit Sub
        End If
    Next

    Err.Raise vbObjectError + 3, MODULE_NAME & "." & PROCEDURE_NAME, "Key not found: " & p_Key
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Sub

Public Sub Remove(ByVal p_Key As String)
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "Remove"

    Dim i As Long
    For i = 1 To m_colItems.Count
        If m_colItems(i).Key = p_Key Then
            m_colItems.Remove i
            Exit Sub
        End If
    Next

    Err.Raise vbObjectError + 4, MODULE_NAME & "." & PROCEDURE_NAME, "Key not found: " & p_Key
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Sub

Public Function Exists(ByVal p_Key As String) As Boolean
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "Exists"

    Dim objKeyValuePair As clsKeyValuePair
    For Each objKeyValuePair In m_colItems
        If objKeyValuePair.Key = p_Key Then
            Exists = True
            Exit Function
        End If
    Next

    Exists = False
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

Public Property Get Count() As Long
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "Count"

    Count = m_colItems.Count
    Exit Property
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Property

Public Sub Clear()
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "Clear"

    Set m_colItems = New Collection
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Sub

Public Function Keys() As Collection
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "Keys"

    Dim colResult As Collection
    Set colResult = New Collection
    Dim objKeyValuePair As clsKeyValuePair
    For Each objKeyValuePair In m_colItems
        colResult.Add objKeyValuePair.Key
    Next

    Set Keys = colResult
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

Public Function Values() As Collection
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "Values"

    Dim colResult As Collection
    Set colResult = New Collection
    Dim objKeyValuePair As clsKeyValuePair
    For Each objKeyValuePair In m_colItems
        colResult.Add objKeyValuePair.Value
    Next

    Set Values = colResult
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

