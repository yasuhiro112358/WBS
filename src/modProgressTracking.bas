Attribute VB_Name = "modProgressTracking"
Sub UpdateProgress(ByVal taskId As Integer)
    Dim ws As Worksheet
    Set ws = shtWBS
    
    'Dim taskId As Integer
    'taskId = InputBox("Enter Task ID to update progress:")
    
    ws.Cells(taskId + 1, 12).value = ws.Cells(taskId + 1, 10).value - ws.Cells(taskId + 1, 11).value ' Remaining Work Hours
    ws.Cells(taskId + 1, 5).value = (ws.Cells(taskId + 1, 11).value / ws.Cells(taskId + 1, 10).value) * 100 ' Progress %

    'MsgBox "Progress updated!", vbInformation
    Debug.Print "Progress has been updated on row # " & taskId
End Sub

