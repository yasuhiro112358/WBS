Attribute VB_Name = "modLogger"
Option Explicit

Sub LogMessage(p_message As String, Optional p_Level As String = "INFO")
    If ENABLE_LOG Then
        Debug.Print "[" & Format(Now, "yyyy-mm-dd HH:MM:SS") & "] [" & p_Level & "] " & p_message
    End If
End Sub

Sub LogError()
    If Err.Number <> 0 Then
        Dim ErrorNumber As Long
        ErrorNumber = GetFriendlyErrorNumber()
    
        LogMessage "Number: " & ErrorNumber, "ERROR"
        LogMessage "Source: " & Err.source, "ERROR"
        LogMessage "Description: " & Err.Description, "ERROR"
        
        Err.Clear
    End If
End Sub

