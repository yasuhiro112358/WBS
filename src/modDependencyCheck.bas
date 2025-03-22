Attribute VB_Name = "modDependencyCheck"
Function IsPredecessorCompleted(taskRow As Integer) As Boolean
    Dim ws As Worksheet
    Set ws = WBSData

    Dim predecessorRow As Integer
    predecessorRow = ws.Cells(taskRow, 15).value

    If ws.Cells(predecessorRow, 14).value = "" Then ' Predecessor task is incomplete
        IsPredecessorCompleted = False
    Else
        IsPredecessorCompleted = True
    End If
End Function

