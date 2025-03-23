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
    Application.EnableEvents = False
     
    If Not Intersect(Target, Me.Range("A1:P256")) Is Nothing Then
        Debug.Print "Target.Address:" & Target.Address
        
        Dim rowChanged As String
        rowChanged = Target.Row
        Debug.Print "rowChanged: " & rowChanged
        
        'UpdateProgress rowChanged
        
    End If
      
    Application.EnableEvents = True
End Sub

