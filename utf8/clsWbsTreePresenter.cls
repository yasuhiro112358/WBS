VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWbsTreePresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const INDENT_WIDTH As Long = 8
Private Const LINE_PREFIX As String = "* "

Private m_objTree As clsWbsTree

Private Sub Class_Initialize()
    Set m_objTree = New clsWbsTree
End Sub

Public Sub InjectTree(ByVal p_objTree As clsWbsTree)
    Set m_objTree = p_objTree
End Sub

Public Sub PrintToDebugWindow()

End Sub

Public Sub ExportToSheet(p_Sheet As Worksheet)
    On Error GoTo ErrorHandler

    With p_Sheet
        .Cells.ClearContents
        
        ' .Cells(1, 1).Value = "Level"
        ' .Cells(1, 2).Value = "Task ID"
        .Cells(1, 1).Value = "Task Name"
    End With

    Dim RowIndex As Long
    RowIndex = 2

    Call ExportNodeRecursive(m_objTree.Root, 0, RowIndex, p_Sheet)
    
    Exit Sub
    
ErrorHandler:
    Call HandleError("clsWbsTreePresenter", "ExportToSheet")

End Sub

Public Sub ExportToCsv(p_FilePath As String)

End Sub

Private Sub ExportNodeRecursive(p_objNode As clsWbsNode, ByVal p_Level As Long, ByRef p_RowIndex As Long, p_Sheet As Worksheet)
    On Error GoTo ErrorHandler
    
    If p_objNode Is Nothing Then
        Exit Sub
    End If

    Dim Indent As String
    Indent = String(p_Level * INDENT_WIDTH, " ")
    
    Dim DisplayName As String
    DisplayName = Indent & LINE_PREFIX & p_objNode.Name

    With p_Sheet
'         .Cells(p_RowIndex, 1).Value = p_Level
'         .Cells(p_RowIndex, 2).Value = p_objNode.Id
        .Cells(p_RowIndex, 1).Value = DisplayName
    End With

    p_RowIndex = p_RowIndex + 1
    
    Dim colChildren As Collection
    Set colChildren = m_objTree.GetChildren(p_objNode.Id)

    Dim objChild As clsWbsNode
    For Each objChild In colChildren
        Call ExportNodeRecursive(objChild, p_Level + 1, p_RowIndex, p_Sheet)
    Next objChild

    Exit Sub
    
ErrorHandler:
    Call HandleError("clsWbsTreePresenter", "ExportNodeRecursive")

End Sub





