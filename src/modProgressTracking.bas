Attribute VB_Name = "modProgressTracking"
Sub UpdateProgress()
    Dim ws As Worksheet
    Set ws = WBSData
    
    Dim rowNum As Integer
    rowNum = InputBox("Enter Task ID to update progress:")
    
    ws.Cells(rowNum + 1, 12).value = ws.Cells(rowNum + 1, 10).value - ws.Cells(rowNum + 1, 11).value ' Remaining Work Hours
    ws.Cells(rowNum + 1, 5).value = (ws.Cells(rowNum + 1, 11).value / ws.Cells(rowNum + 1, 10).value) * 100 ' Progress %

    MsgBox "Progress updated!", vbInformation
End Sub

