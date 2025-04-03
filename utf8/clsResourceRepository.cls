VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsResourceRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME As String = "clsResourceRepository"

Private m_objWorksheet As Worksheet

Private Sub Class_Initialize()
    Set m_objWorksheet = shtResources
End Sub

Public Function FindAll() As Collection
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "FindAll"

    Dim colResources As Collection
    Set colResources = New Collection

    Dim i As Long
    Dim LastRow As Long
    LastRow = m_objWorksheet.Cells(m_objWorksheet.Rows.Count, 1).End(xlUp).Row

    Dim objResource As clsResource
    For i = 2 To LastRow
        Set objResource = New clsResource
        objResource.Id = CStr(m_objWorksheet.Cells(i, 1).Value)
        objResource.Name = CStr(m_objWorksheet.Cells(i, 2).Value)
        
        colResources.Add objResource
    Next i

    Set FindAll = colResources
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

'
' Return `Nothing` if it is not found
'
Public Function Find(p_Id As String) As clsResource
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "Find"

    Dim colResources As Collection
    Set colResources = Me.FindAll()

    Dim objResource As clsResource
    For Each objResource In colResources
        If objResource.Id = p_Id Then
            Set Find = objResource
            Exit Function
        End If
    Next

    Set Find = Nothing
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

