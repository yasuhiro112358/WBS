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

Private Const HEADER_ROW As Long = 4
Private Const DATA_START_ROW As Long = 5

' Siemens Cyan
Private HEADER_BG_COLOR As Long
Private HEADER_FONT_COLOR As Long
Private BORDER_COLOR As Long

Private m_objWorksheet As Worksheet

Private Sub Class_Initialize()
    HEADER_BG_COLOR = RGB(0, 153, 153)
    HEADER_FONT_COLOR = RGB(255, 255, 255)
    BORDER_COLOR = RGB(64, 64, 64)
End Sub

Public Sub Init(p_objWorksheet As Worksheet)
    Set m_objWorksheet = p_objWorksheet
End Sub

Public Sub Clear()
    Dim LastRow As Long
    LastRow = m_objWorksheet.Cells(m_objWorksheet.Rows.Count, 1).End(xlUp).Row

    If LastRow > 1 Then
        m_objWorksheet.Rows(DATA_START_ROW & ":" & LastRow).ClearContents
    End If
End Sub

Public Sub Render(p_colWbsViewItems As Collection)
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "Render"

    Call DrawHeader
    
    With m_objWorksheet.Range(m_objWorksheet.Cells(DATA_START_ROW, 1), m_objWorksheet.Cells(m_objWorksheet.Rows.Count, 1))
        .HorizontalAlignment = xlLeft
    End With
    With m_objWorksheet.Range(m_objWorksheet.Cells(DATA_START_ROW, 15), m_objWorksheet.Cells(m_objWorksheet.Rows.Count, 15))
        .HorizontalAlignment = xlLeft
    End With

    Dim i As Long
    i = DATA_START_ROW
    Dim objWbsViewItem As clsWbsViewItem
    For Each objWbsViewItem In p_colWbsViewItems
        With m_objWorksheet
'             .Cells(i, 1).Value = i - DATA_START_ROW + 1
            .Cells(i, 1).Value = objWbsViewItem.WbsCode
'            .Cells(i, 1).Formula = "=""" & objWbsViewItem.WbsCode & """"
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
            .Cells(i, 15).Value = objWbsViewItem.PredecessorWbsCode
            ' .Cells(i, 15).Formula = "=""" & objWbsViewItem.PredecessorWbsCode & """"
            .Cells(i, 16).Value = objWbsViewItem.PredecessorName
            .Cells(i, 17).Value = objWbsViewItem.ResourceName
            .Cells(i, 18).Value = objWbsViewItem.TaskId
            .Cells(i, 19).Value = objWbsViewItem.PredecessorId
            .Cells(i, 20).Value = objWbsViewItem.ResourceId
        End With
        i = i + 1
    Next
    
    Dim DataEndRow As Long
    DataEndRow = i - 1
    Dim DataRange As Range
    Set DataRange = m_objWorksheet.Range(m_objWorksheet.Cells(DATA_START_ROW, 1), m_objWorksheet.Cells(DataEndRow, 20))
    With DataRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = BORDER_COLOR
    End With
        
    With m_objWorksheet
        .Range(.Cells(DATA_START_ROW, 3), .Cells(DataEndRow, 3)).NumberFormat = "mm/dd"
        .Range(.Cells(DATA_START_ROW, 4), .Cells(DataEndRow, 4)).NumberFormat = "mm/dd"
        .Range(.Cells(DATA_START_ROW, 7), .Cells(DataEndRow, 7)).NumberFormat = "mm/dd"
        .Range(.Cells(DATA_START_ROW, 8), .Cells(DataEndRow, 8)).NumberFormat = "mm/dd"
        .Range(.Cells(DATA_START_ROW, 13), .Cells(DataEndRow, 13)).NumberFormat = "mm/dd"
        .Range(.Cells(DATA_START_ROW, 14), .Cells(DataEndRow, 14)).NumberFormat = "mm/dd"
    
        .Columns("A:T").AutoFit
        
        If .Parent.Windows(1).FreezePanes Then
            .Parent.Windows(1).FreezePanes = False
        End If
'        .Range("C5").Select
        .Rows(5).Select
        .Parent.Windows(1).FreezePanes = True
    End With
    
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

Private Sub DrawHeader()
    With m_objWorksheet
'         .Cells(HEADER_ROW, 1).Value = "#"
        .Cells(HEADER_ROW, 1).Value = "WBS Code"
        .Cells(HEADER_ROW, 2).Value = "Task Name"
        .Cells(HEADER_ROW, 3).Value = "Start Date"
        .Cells(HEADER_ROW, 4).Value = "End Date"
        .Cells(HEADER_ROW, 5).Value = "Progress [%]"
        .Cells(HEADER_ROW, 6).Value = "Duration"
        .Cells(HEADER_ROW, 7).Value = "Baseline" & vbLf & "Start Date"
        .Cells(HEADER_ROW, 8).Value = "Baseline" & vbLf & "End Date"
        .Cells(HEADER_ROW, 9).Value = "Baseline" & vbLf & "Work Hours"
        .Cells(HEADER_ROW, 10).Value = "Assigned" & vbLf & "Work Hours"
        .Cells(HEADER_ROW, 11).Value = "Actual" & vbLf & "Work Hours"
        .Cells(HEADER_ROW, 12).Value = "Remaining" & vbLf & "Work Hours"
        .Cells(HEADER_ROW, 13).Value = "Actual" & vbLf & "Start Date"
        .Cells(HEADER_ROW, 14).Value = "Actual" & vbLf & "End Date"
        .Cells(HEADER_ROW, 15).Value = "Predecessor" & vbLf & "WBS Code"
        .Cells(HEADER_ROW, 16).Value = "Predecessor" & vbLf & "Task Name"
        .Cells(HEADER_ROW, 17).Value = "Resource Name"
        .Cells(HEADER_ROW, 18).Value = "Task ID"
        .Cells(HEADER_ROW, 19).Value = "Predecessor" & vbLf & "Task ID"
        .Cells(HEADER_ROW, 20).Value = "Resource ID"
                
        With .Cells(HEADER_ROW, 1).Resize(1, 20)
            .Font.Bold = True
            .Font.Color = HEADER_FONT_COLOR
            .Interior.Color = HEADER_BG_COLOR
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .Orientation = xlUpward
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = BORDER_COLOR
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlDouble
                .Weight = xlThick
            End With
        End With
    End With
End Sub

