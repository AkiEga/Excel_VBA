Attribute VB_Name = "CSV"
Option Explicit



Function SetCSVFolder() As String
    Dim FolderNameCell As Range
    Set FolderNameCell = Worksheets(1).Cells(1, 2)
    'フォルダを選択する
    'MyFol = Worksheets(1).Range("D2").Value
    'MyFol = FolderNameCell.Value
    'フォルダが存在しないならば再度フォルダ選択
    If FolderNameCell.Value = "" Then
        Call SelectFolder(FolderNameCell)
    End If
    SetCSVFolder = FolderNameCell.Value
End Function

Sub SelectFolder(FolderNameCell As Range)
    Dim MyObj As Object
    Dim MyFol, MyFnm As String

    'フォルダを選択する
    Set MyObj = CreateObject("Shell.Application") _
        .BrowseForFolder(0, "SelectFolder", 0, ActiveWorkbook.Path)
    '選択なければ処理を抜ける
    If MyObj Is Nothing Then Exit Sub
    MyFol = MyObj.self.Path & "\"
    
    '選択フォルダ名をセルに書き込み
    FolderNameCell.Value = MyFol
    MsgBox MyFol & "を処理します。"
End Sub


