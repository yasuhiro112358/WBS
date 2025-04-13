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
    
    Dim objWbsApplication As clsWbsApplication
    Set objWbsApplication = New clsWbsApplication
    Call objWbsApplication.Init
    
    Call objWbsApplication.UpdateTask(p_objTarget)
    
    Call objWbsApplication.RefreshWbsView
        
Cleanup:
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbExclamation
    Resume Cleanup
End Sub
