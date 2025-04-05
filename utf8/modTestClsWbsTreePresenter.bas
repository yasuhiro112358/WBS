Attribute VB_Name = "modTestClsWbsTreePresenter"
Option Explicit

Sub TestClsWbsTreePresenter()
    On Error GoTo ErrorHandler

    Dim objDbSheet As Worksheet
    Set objDbSheet = shtDb
    
    Dim objStorage As clsWbsTreeStorage
    Set objStorage = New clsWbsTreeStorage
    
    Dim objTree As clsWbsTree
    Set objTree = objStorage.LoadFromSheet(objDbSheet)
    
    Call LogMessage("Complete loading.")
     
    Dim objOutputSheet As Worksheet
    Set objOutputSheet = shtTest2
     
    Dim objPresenter As clsWbsTreePresenter
    Set objPresenter = New clsWbsTreePresenter
    Call objPresenter.InjectTree(objTree)
    Call objPresenter.ExportToSheet(objOutputSheet)

    Call LogMessage("Complete exporting to sheet.")
    
    Exit Sub
    
ErrorHandler:
    Call HandleError("modTestClsWbsTreePresenter", "TestClsWbsTreePresenter")
    
End Sub

