Option Explicit

Sub ExportVBAModules()
    Dim vbComp As Object
    Dim exportPath As String
    
    exportPath = ThisWorkbook.Path & "\src\"
    
    ' フォルダが存在しない場合は作成
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If
    
    ' 各モジュールをエクスポート
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1 ' 標準モジュール
                vbComp.Export exportPath & vbComp.Name & ".bas"
            Case 2 ' クラスモジュール
                vbComp.Export exportPath & vbComp.Name & ".cls"
            Case 3 ' ユーザーフォーム
                vbComp.Export exportPath & vbComp.Name & ".frm"
        End Select
    Next vbComp
    
    MsgBox "VBAコードをエクスポートしました！", vbInformation
End Sub

Sub ImportVBAModules()
    Dim vbComp As Object
    Dim importPath As String
    Dim fileName As String
    Dim moduleName As String
    
    importPath = ThisWorkbook.Path & "\src\"
    
    ' .bas ファイルのインポート
    fileName = Dir(importPath & "*.bas")
    Do While fileName <> ""
        moduleName = Left(fileName, InStrRev(fileName, ".") - 1)
        
        ' 既存のモジュールを削除（エラーが発生した場合は無視）
        On Error Resume Next
        Set vbComp = ThisWorkbook.VBProject.VBComponents(moduleName)
        If Not vbComp Is Nothing Then
            ThisWorkbook.VBProject.VBComponents.Remove vbComp
        End If
        Err.Clear
        On Error GoTo 0

        ' 新しいモジュールをインポート
        ThisWorkbook.VBProject.VBComponents.Import importPath & fileName
        fileName = Dir
    Loop

    ' .cls ファイルのインポート
    fileName = Dir(importPath & "*.cls")
    Do While fileName <> ""
        moduleName = Left(fileName, InStrRev(fileName, ".") - 1)
        
        ' 既存のクラスモジュールを削除
        On Error Resume Next
        Set vbComp = ThisWorkbook.VBProject.VBComponents(moduleName)
        If Not vbComp Is Nothing Then
            ThisWorkbook.VBProject.VBComponents.Remove vbComp
        End If
        Err.Clear
        On Error GoTo 0

        ' インポート
        ThisWorkbook.VBProject.VBComponents.Import importPath & fileName
        fileName = Dir
    Loop

    ' .frm ファイルのインポート（ユーザーフォーム）
    fileName = Dir(importPath & "*.frm")
    Do While fileName <> ""
        moduleName = Left(fileName, InStrRev(fileName, ".") - 1)
        
        ' 既存のユーザーフォームを削除
        On Error Resume Next
        Set vbComp = ThisWorkbook.VBProject.VBComponents(moduleName)
        If Not vbComp Is Nothing Then
            ThisWorkbook.VBProject.VBComponents.Remove vbComp
        End If
        Err.Clear
        On Error GoTo 0

        ' インポート
        ThisWorkbook.VBProject.VBComponents.Import importPath & fileName
        fileName = Dir
    Loop

    MsgBox "VBAコードをインポートしました！", vbInformation
End Sub
