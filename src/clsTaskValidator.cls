VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTaskValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_colErrors As Collection

Public Function Validate(p_Task As clsTask) As Boolean
    Set m_colErrors = New Collection
    
    If Trim(p_Task.WbsCode) = "" Then
        m_colErrors.Add "WBS Code is required."
    End If

    If Len(Trim(p_Task.Name)) = 0 Then
        m_colErrors.Add "Task name is required."
    End If

    If p_Task.Progress < 0 Or p_Task.Progress > 100 Then
        m_colErrors.Add "Progress must be between 0 and 100."
    End If

    If p_Task.StartDate > 0 And p_Task.EndDate > 0 Then
        If p_Task.EndDate < p_Task.StartDate Then
            m_colErrors.Add "End date cannot be before start date."
        End If
    End If

    Validate = (m_colErrors.Count = 0)
End Function

Public Function GetErrors() As Collection
    Set GetErrors = m_colErrors
End Function

