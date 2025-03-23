VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtWBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    LogMessage "[Worksheet_Change] Running..."
    Application.EnableEvents = False
     
    If Not Intersect(Target, Me.Range("A1:P256")) Is Nothing Then
        LogMessage "[Worksheet_Change] Target.Address:" & Target.Address
        
        Dim rowChanged As String
        rowChanged = Target.row
        ' Debug.Print "rowChanged: " & rowChanged
                
        Dim TaskIdChanged As String
        TaskIdChanged = Me.Cells(rowChanged, 1).value
        If Not IsEmpty(TaskIdChanged) Then
            LogMessage "[Worksheet_Change] Task ID: " & TaskIdChanged
        Else
            LogMessage "[Worksheet_Change] Task ID is empty or does not exist."
        End If
        
        ' AutoUpdateTask TaskIdChanged
        
        
    End If
      
    Application.EnableEvents = True
    LogMessage "[Worksheet_Change] Completed"
End Sub

