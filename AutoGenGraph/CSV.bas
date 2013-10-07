Attribute VB_Name = "CSV"
Option Explicit



Function SetCSVFolder() As String
    Dim FolderNameCell As Range
    Set FolderNameCell = Worksheets(1).Cells(1, 2)
    '�t�H���_��I������
    'MyFol = Worksheets(1).Range("D2").Value
    'MyFol = FolderNameCell.Value
    '�t�H���_�����݂��Ȃ��Ȃ�΍ēx�t�H���_�I��
    If FolderNameCell.Value = "" Then
        Call SelectFolder(FolderNameCell)
    End If
    SetCSVFolder = FolderNameCell.Value
End Function

Sub SelectFolder(FolderNameCell As Range)
    Dim MyObj As Object
    Dim MyFol, MyFnm As String

    '�t�H���_��I������
    Set MyObj = CreateObject("Shell.Application") _
        .BrowseForFolder(0, "SelectFolder", 0, ActiveWorkbook.Path)
    '�I���Ȃ���Ώ����𔲂���
    If MyObj Is Nothing Then Exit Sub
    MyFol = MyObj.self.Path & "\"
    
    '�I���t�H���_�����Z���ɏ�������
    FolderNameCell.Value = MyFol
    MsgBox MyFol & "���������܂��B"
End Sub


