VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWbsTreeStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const HEADER_ROW As Long = 1
Private Const COL_TASK_ID As Long = 1
Private Const COL_PARENT_ID As Long = 2
Private Const COL_NAME As Long = 3

Public Sub SaveToSheet(ByVal p_objSheet As Worksheet, ByVal p_objTree As clsWbsTree)
    If p_objTree Is Nothing Or p_objSheet Is Nothing Then
        Exit Sub
    End If

    p_objSheet.Cells.ClearContents
    
    With p_objSheet
        .Cells(HEADER_ROW, COL_TASK_ID).Value = "Task ID"
        .Cells(HEADER_ROW, COL_PARENT_ID).Value = "Parent ID"
        .Cells(HEADER_ROW, COL_NAME).Value = "Task Name"
    End With

    Dim RowIndex As Long
    RowIndex = HEADER_ROW + 1
    
    Call SaveNodeRecursive(p_objTree.Root, p_objTree, RowIndex, p_objSheet)
End Sub

Private Sub SaveNodeRecursive(ByVal p_objNode As clsWbsNode, ByVal p_objTree As clsWbsTree, ByRef p_RowIndex As Long, ByVal p_objSheet As Worksheet)
    If p_objNode Is Nothing Then
        Exit Sub
    End If

    Dim objParent As clsWbsNode
    Set objParent = p_objTree.GetParent(p_objNode.Id)

    With p_objSheet
        .Cells(p_RowIndex, COL_TASK_ID).Value = p_objNode.Id
        
        If objParent Is Nothing Then
            .Cells(p_RowIndex, COL_PARENT_ID).Value = ""
        Else
            .Cells(p_RowIndex, COL_PARENT_ID).Value = objParent.Id
        End If
        
        .Cells(p_RowIndex, COL_NAME).Value = p_objNode.Name
    End With

    p_RowIndex = p_RowIndex + 1

    Dim colChildren As Collection
    Set colChildren = p_objTree.GetChildren(p_objNode.Id)

    Dim objChild As clsWbsNode
    For Each objChild In colChildren
        Call SaveNodeRecursive(objChild, p_objTree, p_RowIndex, p_objSheet)
    Next objChild
End Sub

Public Function LoadFromSheet(ByVal p_Sheet As Worksheet) As clsWbsTree
    ' ??????
    ' ??????(Task ID / Parent ID ????????)
End Function

