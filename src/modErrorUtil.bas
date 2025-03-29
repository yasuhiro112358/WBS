Attribute VB_Name = "modErrorUtil"
Option Explicit

Public Sub RaiseError(ByVal p_Number As Long, _
                      Optional ByVal p_Source As String = "", _
                      Optional ByVal p_Description As String = "")
                      
    Err.Raise vbObjectError + p_Number, p_Source, p_Description
End Sub

Public Function GetFriendlyErrorNumber() As Long
    If IsCustomError Then
        GetFriendlyErrorNumber = Err.Number - vbObjectError
    Else
        GetFriendlyErrorNumber = Err.Number
    End If
End Function

Public Function IsCustomError() As Boolean
    IsCustomError = (Err.Number >= vbObjectError)
End Function

