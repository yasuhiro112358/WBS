VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWbsPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME As String = "clsWbsPresenter"

Private Const INDENT_WIDTH As Long = 8
'Private Const LINE_PREFIX As String = "* "
Private Const LINE_PREFIX As String = ""

Private m_objTaskService As clsTaskService

Public Sub Init(p_objTaskService As clsTaskService)
    Set m_objTaskService = p_objTaskService
End Sub

Public Function BuildWbsViewItems(p_colTasks As Collection, p_colResources As Collection) As Collection
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "BuildViewItems"
    
    Dim dictTaskIdToLineNo As clsDictionary
    Set dictTaskIdToLineNo = New clsDictionary
    
    Dim dictTaskIdToTask As clsDictionary
    Set dictTaskIdToTask = New clsDictionary

    Dim dictWbsCodeToTask As clsDictionary
    Set dictWbsCodeToTask = New clsDictionary
    
    Dim LineNo As Long
    LineNo = 1
    Dim objTask As clsTask
    For Each objTask In p_colTasks
        Call dictTaskIdToLineNo.Add(objTask.Id, LineNo)
        
        Call dictTaskIdToTask.Add(objTask.Id, objTask)
        
        If objTask.WbsCode <> "" Then
            Call dictWbsCodeToTask.Add(objTask.WbsCode, objTask)
        End If
        LineNo = LineNo + 1
    Next


    Dim colWbsViewItems As Collection
    Set colWbsViewItems = New Collection
    ' Dim objTask As clsTask
    Dim objWbsViewItem As clsWbsViewItem
    
'     Dim objPredecessorTask As clsTask
    Dim objResource As clsResource
    
    LineNo = 1
    For Each objTask In p_colTasks
        Set objWbsViewItem = New clsWbsViewItem
        With objWbsViewItem
            .TaskId = objTask.Id
            .WbsCode = objTask.WbsCode
            .PredecessorId = objTask.PredecessorId
            .ResourceId = objTask.ResourceId
            
            .TaskLineNo = LineNo
            If dictTaskIdToLineNo.Exists(objTask.ParentId) Then
                .ParentLineNo = dictTaskIdToLineNo.Item(objTask.ParentId)
            End If
            If dictTaskIdToLineNo.Exists(objTask.PredecessorId) Then
                .PredecessorLineNo = dictTaskIdToLineNo.Item(objTask.PredecessorId)
            End If

            .Name = objTask.Name
            .Level = m_objTaskService.GetLevel(objTask)
            .DisplayName = CreateDisplayName(.Name, .Level)
            .StartDate = objTask.StartDate
            .EndDate = objTask.EndDate
            .Progress = objTask.Progress
            .Duration = objTask.Duration
            .BaselineStartDate = objTask.BaselineStartDate
            .BaselineEndDate = objTask.BaselineEndDate
            .BaselineWorkHours = objTask.BaselineWorkHours
            .AssignedWorkHours = objTask.AssignedWorkHours
            .ActualWorkHours = objTask.ActualWorkHours
            
            .RemainingWorkHours = m_objTaskService.CalcRemainingWorkHours(objTask)
            
            .ActualStartDate = objTask.ActualStartDate
            .ActualEndDate = objTask.ActualEndDate
            
            .StatusLabel = m_objTaskService.GetStatus(objTask)
            
'            If dictTaskIdToTask.Exists(objTask.PredecessorId) Then
'                .PredecessorName = dictTaskIdToTask.Item(objTask.PredecessorId).Name
'            End If
            If objTask.PredecessorId <> "" And dictTaskIdToTask.Exists(objTask.PredecessorId) Then
                Dim objPredecessorTask As clsTask
                Set objPredecessorTask = dictTaskIdToTask.Item(objTask.PredecessorId)
                .PredecessorWbsCode = objPredecessorTask.WbsCode
                .PredecessorName = objPredecessorTask.Name
            End If
            
        End With
        
'        Set objPredecessorTask = FindTaskById(p_colTasks, objTask.PredecessorId)
'        If Not objPredecessorTask Is Nothing Then
'            objWbsViewItem.PredecessorName = objPredecessorTask.Name
'        Else
'            objWbsViewItem.PredecessorName = ""
'        End If

        Set objResource = FindResourceById(p_colResources, objTask.ResourceId)
        If Not objResource Is Nothing Then
            objWbsViewItem.ResourceName = objResource.Name
        Else
            objWbsViewItem.ResourceName = "(Unassigned)"
        End If

        colWbsViewItems.Add objWbsViewItem
        
        LineNo = LineNo + 1
    Next objTask

'    Set BuildWbsViewItems = colWbsViewItems
'    Set BuildWbsViewItems = SortCollectionByProperty(colWbsViewItems, "WbsCode")
    Set BuildWbsViewItems = SortWbsViewItemsByWbsCode(colWbsViewItems)
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

Private Function CreateDisplayName(p_Name As String, p_Level As Long) As String
    CreateDisplayName = String((p_Level - 1) * INDENT_WIDTH, " ") & LINE_PREFIX & p_Name
End Function

Private Function FindTaskById(p_colTasks As Collection, p_Id As String) As clsTask
    Dim objTask As clsTask
    For Each objTask In p_colTasks
        If objTask.Id = p_Id Then
            Set FindTaskById = objTask
            Exit Function
        End If
    Next
    
    Set FindTaskById = Nothing
End Function

Private Function FindResourceById(p_colResources As Collection, p_Id As String) As clsResource
    Dim objResource As clsResource
    For Each objResource In p_colResources
        If objResource.Id = p_Id Then
            Set FindResourceById = objResource
            Exit Function
        End If
    Next
    
    Set FindResourceById = Nothing
End Function

'
' Return an array made from splitted WBS code.
' Type: Array(Value as Long) as Variant
'
Private Function ParseWbsCode(p_Code As String) As Variant
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "ParseWbsCode"

    Dim CleanCode As String
    If UCase(Left(p_Code, 4)) = "WBS-" Then
        CleanCode = Mid(p_Code, 5)
    Else
        Err.Raise Err.Number + 1000, , "WBS code on wrong format found."
    End If

    Dim arrParts As Variant
    arrParts = Split(CleanCode, ".")
    
    Dim i As Long
    For i = LBound(arrParts) To UBound(arrParts)
        If IsNumeric(arrParts(i)) Then
            arrParts(i) = CLng(arrParts(i))
        Else
            arrParts(i) = 0
        End If
    Next i
    
    ParseWbsCode = arrParts
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

'
' Retern -1 if p_Code1 is less than p_Code2
' Retern  0 if p_Code1 is equal to p_Code2
' Retern  1 if p_Code1 is greater than p_Code2
'
Private Function CompareWbsCodes(p_Code1 As String, p_Code2 As String) As Integer
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "CompareWbsCodes"
        
    Dim arr1 As Variant
    arr1 = ParseWbsCode(p_Code1)
    Dim arr2 As Variant
    arr2 = ParseWbsCode(p_Code2)
    
    Dim Len1 As Long
    Len1 = UBound(arr1)
    Dim Len2 As Long
    Len2 = UBound(arr2)
    
    Dim MaxLen As Long
    MaxLen = Application.WorksheetFunction.Max(Len1, Len2)
    
    Dim i As Long
    For i = 0 To MaxLen
        Dim val1 As Long
        If i <= Len1 Then
            val1 = arr1(i)
        Else
            val1 = 0
        End If
        
        Dim val2 As Long
        If i <= Len2 Then
            val2 = arr2(i)
        Else
            val2 = 0
        End If
        
        If val1 < val2 Then
            CompareWbsCodes = -1
            Exit Function
        ElseIf val1 > val2 Then
            CompareWbsCodes = 1
            Exit Function
        End If
    Next
    
    CompareWbsCodes = 0
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function

Private Function SortWbsViewItemsByWbsCode(p_colWbsViewItems As Collection) As Collection
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "SortWbsViewItemsByWbsCode"
    
    Dim arr() As clsWbsViewItem
    ReDim arr(1 To p_colWbsViewItems.Count)
    
    Dim i As Long
    For i = 1 To p_colWbsViewItems.Count
        Set arr(i) = p_colWbsViewItems(i)
    Next
    
    For i = 1 To UBound(arr) - 1
        Dim j As Long
        For j = i + 1 To UBound(arr)
            If CompareWbsCodes(arr(i).WbsCode, arr(j).WbsCode) > 0 Then
                Dim tmp As clsWbsViewItem
                Set tmp = arr(i)
                Set arr(i) = arr(j)
                Set arr(j) = tmp
            End If
        Next j
    Next i

    Dim colSorted As Collection
    Set colSorted = New Collection
    For i = 1 To UBound(arr)
        colSorted.Add arr(i)
    Next i
    
    Set SortWbsViewItemsByWbsCode = colSorted
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function
