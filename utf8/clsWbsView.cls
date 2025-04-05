VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWbsView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME As String = "clsWbsView"

Private m_objWorksheet As Worksheet

Public Sub Init(p_objWorksheet As Worksheet)
    Set m_objWorksheet = p_objWorksheet
End Sub

Public Sub Clear()
    Dim LastRow As Long
    LastRow = m_objWorksheet.Cells(m_objWorksheet.Rows.Count, 1).End(xlUp).Row

    If LastRow > 1 Then
        m_objWorksheet.Rows("2:" & LastRow).ClearContents
    End If
End Sub

Public Sub Render(p_colWbsViewItems As Collection)
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "Render"

    Dim i As Long
    i = 2

    Dim objWbsViewItem As clsWbsViewItem
    For Each objWbsViewItem In p_colWbsViewItems
        With m_objWorksheet
            .Cells(i, 1).Value = objWbsViewItem.Id
            .Cells(i, 2).Value = objWbsViewItem.DisplayName
            .Cells(i, 3).Value = FormatDate(objWbsViewItem.StartDate)
            .Cells(i, 4).Value = FormatDate(objWbsViewItem.EndDate)
            .Cells(i, 5).Value = objWbsViewItem.Progress
            .Cells(i, 6).Value = objWbsViewItem.Duration
            .Cells(i, 7).Value = FormatDate(objWbsViewItem.BaselineStartDate)
            .Cells(i, 8).Value = FormatDate(objWbsViewItem.BaselineEndDate)
            .Cells(i, 9).Value = objWbsViewItem.BaselineWorkHours
            .Cells(i, 10).Value = objWbsViewItem.AssignedWorkHours
            .Cells(i, 11).Value = objWbsViewItem.ActualWorkHours
            .Cells(i, 12).Value = objWbsViewItem.RemainingWorkHours
            .Cells(i, 13).Value = FormatDate(objWbsViewItem.ActualStartDate)
            .Cells(i, 14).Value = FormatDate(objWbsViewItem.ActualEndDate)
            .Cells(i, 15).Value = objWbsViewItem.PredecessorId
            .Cells(i, 16).Value = objWbsViewItem.ResourceName
        End With
        i = i + 1
    Next

    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Sub

Private Function FormatDate(p_Date As Date) As Variant
    If p_Date = 0 Then
        FormatDate = ""
    Else
        FormatDate = p_Date
    End If
End Function

