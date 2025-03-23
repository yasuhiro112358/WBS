Attribute VB_Name = "modLogger"
Option Explicit

Const ENABLE_LOG As Boolean = True

Sub LogMessage(message As String, Optional level As String = "INFO")
    If ENABLE_LOG Then
        Debug.Print "[" & Format(Now, "yyyy-mm-dd HH:MM:SS") & "] [" & level & "] " & message
    End If
End Sub

