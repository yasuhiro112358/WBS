VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCustomDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private keyCollection As Collection
Private valueCollection As Collection

' Initialize Class
Private Sub Class_Initialize()
    Set keyCollection = New Collection
    Set valueCollection = New Collection
    ' Debug.Print "CustomDictionary is instantiated"
End Sub

' Add key and value
Public Sub Add(ByVal Key As String, ByVal Value As Variant)
    Dim i As Integer
    i = GetIndex(Key)
        
    If i > 0 Then
        ' Update value when the key exists
        valueCollection.Remove i
        valueCollection.Add Value, Before:=i
    Else
        ' Create new value
        keyCollection.Add Key
        valueCollection.Add Value
    End If
End Sub

' Check if a key exists
Public Function Exists(ByVal Key As String) As Boolean
    'On Error Resume Next
    Exists = (GetIndex(Key) > 0)
    'On Error GoTo 0
End Function

' Get value by key
Public Property Get Item(ByVal Key As String) As Variant
    Dim i As Integer
    i = GetIndex(Key)
    If i > 0 Then
        Item = valueCollection(i)
    Else
        Err.Raise 9, "CustomDictionary", "Key not found"
    End If
End Property

' Set value with key (update available)
Public Property Let Item(ByVal Key As String, ByVal Value As Variant)
    Add Key, Value
End Property

' Delete value by key
Public Sub Remove(ByVal Key As String)
    Dim i As Integer
    i = GetIndex(Key)
    If i > 0 Then
        keyCollection.Remove i
        valueCollection.Remove i
        Debug.Print "Data having key: " & Key & " is Removed"
    Else
        Err.Raise 9, "CustomDictionary", "Key not found"
    End If
End Sub

' Count the number of keys
Public Property Get Count() As Integer
    Count = keyCollection.Count
End Property

' Get all keys
Public Function Keys() As Variant
    Dim arr() As String
    Dim i As Integer
    
    If keyCollection.Count = 0 Then
        Keys = Array() ' Return blank array
        Exit Function
    End If
    
    ReDim arr(1 To keyCollection.Count)
    For i = 1 To keyCollection.Count
        arr(i) = keyCollection(i)
    Next i
    
    Keys = arr
End Function

' Get all values
Public Function Values() As Variant
    Dim arr() As Variant
    Dim i As Integer
    
    If valueCollection.Count = 0 Then
        Values = Array() ' Return blank array
        Exit Function
    End If
    
    ReDim arr(1 To valueCollection.Count)
    For i = 1 To valueCollection.Count
        arr(i) = valueCollection(i)
    Next i
    
    Values = arr
End Function

' Remove all data
Public Sub RemoveAll()
    Set keyCollection = New Collection
    Set valueCollection = New Collection
    Debug.Print "All data is Removed"
End Sub

' Get index by key
Private Function GetIndex(ByVal Key As String) As Integer
    Dim i As Integer
    
    If keyCollection.Count = 0 Then
        GetIndex = 0
        Exit Function
    End If
    
    For i = 1 To keyCollection.Count
        If keyCollection(i) = Key Then
            GetIndex = i
            Exit Function
        End If
    Next i
    
    ' Not found
    GetIndex = 0
End Function


