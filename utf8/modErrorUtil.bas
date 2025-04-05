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

Public Sub HandleError(ByVal ModuleName As String, ByVal ProcedureName As String)
    Dim ErrorMessage As String
    ErrorMessage = vbCrLf & _
                   "Number:      " & Err.Number & vbCrLf & _
                   "Module:      " & ModuleName & vbCrLf & _
                   "Procedure:   " & ProcedureName & vbCrLf & _
                   "Description: " & Err.Description
                 
    ' ErrorMessage = "Num: " & Err.Number & " Src: " & ModuleName & "." & ProcedureName & " Desc: " & Err.Description

    Call LogMessage(ErrorMessage, "ERROR")

    If IS_TEST_MODE Then
        Err.Clear
    Else
        Err.Raise Err.Number, ModuleName & "." & ProcedureName, Err.Description
    End If
End Sub




