Attribute VB_Name = "modCollectionSorter"
Option Explicit

Public Function SortCollectionByProperty(p_colObjects As Collection, p_PropertyName As String) As Collection
    Dim arrObjects() As Object
    ReDim arrObjects(1 To p_colObjects.Count)
    
    Dim i As Long
    i = 1
    Dim obj As Object
    For Each obj In p_colObjects
        Set arrObjects(i) = obj
        i = i + 1
    Next

    Dim j As Long
    Dim objTemp As Object
    For i = LBound(arrObjects) To UBound(arrObjects) - 1
        For j = i + 1 To UBound(arrObjects)
            If GetPropertyValue(arrObjects(i), p_PropertyName) > GetPropertyValue(arrObjects(j), p_PropertyName) Then
                Set objTemp = arrObjects(i)
                Set arrObjects(i) = arrObjects(j)
                Set arrObjects(j) = objTemp
            End If
        Next j
    Next i

    Dim colSorted As Collection
    Set colSorted = New Collection
    For i = LBound(arrObjects) To UBound(arrObjects)
        colSorted.Add arrObjects(i)
    Next

    Set SortCollectionByProperty = colSorted
End Function

Private Function GetPropertyValue(p_obj As Object, p_PropertyName As String) As Variant
    GetPropertyValue = CallByName(p_obj, p_PropertyName, VbGet)
End Function

