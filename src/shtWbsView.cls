VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtWbsView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal p_objTarget As Range)
    On Error GoTo ErrorHandler
    Application.EnableEvents = False

    Const DATA_START_ROW As Long = 5
    Const COL_TASK_ID As Long = 18

    If p_objTarget.Cells.Count > 1 Then
        GoTo Cleanup
    End If
    
    If p_objTarget.Row < DATA_START_ROW Then
        GoTo Cleanup
    End If
    
    Dim TaskId As String
    TaskId = Me.Cells(p_objTarget.Row, COL_TASK_ID).Value
    If TaskId = "" Then
        Err.Raise Err.Number, , "Task ID not found: " & Err.Description
    End If

    Dim objWbsApplication As clsWbsApplication
    Set objWbsApplication = New clsWbsApplication
    Call objWbsApplication.Init
    
    Call objWbsApplication.UpdateTask(p_objTarget.Value, TaskId, p_objTarget.Column)
    
    Call objWbsApplication.RefreshWbsView
    
Cleanup:
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbExclamation
    Resume Cleanup
End Sub
