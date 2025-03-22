Attribute VB_Name = "modTaskManagement"
Sub AddTask()
    Dim ws As Worksheet
    Set ws = WBSData
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    'ws.Cells(lastRow, 1).value = lastRow - 1 ' Task ID
    ws.Cells(lastRow, 1).value = GenerateUUIDv4()
    ws.Cells(lastRow, 2).value = InputBox("Enter Task Name:")
    ws.Cells(lastRow, 7).value = InputBox("Enter Baseline Start Date (yyyy/mm/dd):")
    ws.Cells(lastRow, 8).value = InputBox("Enter Baseline End Date (yyyy/mm/dd):")
    ws.Cells(lastRow, 9).value = InputBox("Enter Baseline Work Hours:")
    
    MsgBox "Task has been added successfully!", vbInformation
End Sub

