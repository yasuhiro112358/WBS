VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME As String = "ModuleName"

Private Sub Class_Initialize()
    ' Constructor
End Sub

Public Function MethodName() As Variant
    On Error GoTo ErrorHandler
    Const PROCEDURE_NAME As String = "ProcedureName"

    ' ... Main procedure ...

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & PROCEDURE_NAME, Err.Description
End Function



