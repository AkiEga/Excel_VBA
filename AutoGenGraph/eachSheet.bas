Attribute VB_Name = "eachSheet"
'�Q�l�T�C�g: http://oshiete.goo.ne.jp/qa/3168255.html?from=recommend
Dim MyObj As Object
Dim MyFol As String
Dim MyFnm As String
Dim MyStr As String
Dim i   As Integer
Dim n   As Long
Dim n1  As Long

Sub EachSheetReadCSV()
    '�������łɃV�[�g������Ă�����ēx��蒼��
    If ActiveWorkbook.Sheets.Count >= 2 Then
        SheetClear
    End If
    '�t�H���_��I������
    MyFol = SetCSVFolder
    MsgBox MyFol & "���������܂��B"
    
    'Dir�֐����g���Ďw��t�H���_��csv�t�@�C������������
    MyFnm = Dir(MyFol & "*.csv")
    i = 0
    Do Until Len(MyFnm) = 0&
        '���ڂ̃t�@�C������荞�񂾂��J�E���g
        i = i + 1
        'ThisWorkbook�ɃV�[�g��ǉ����ď���
        With Sheets.Add(after:=Worksheets(Worksheets.Count))
            '�t�@�C��������V�[�g���𒊏o
            .Name = Mid(MyFnm, 1, Len(MyFnm) - Len(".csv"))
            '�O���f�[�^��荞�݂𗘗p
            Call ReadCSVFile(MyFol, MyFnm)
        End With
        '���̃t�@�C����
        MyFnm = Dir()
    Loop

End Sub

Sub EachSheetGenGraph()
    i = 2
    Do Until i > Sheets.Count
        Debug.Print CStr(i) + "�V�[�g�ڂ̃O���t���쐬���܂�"
        '���ڂ̃t�@�C������荞�񂾂��J�E���g
        '�O���f�[�^��荞�݂𗘗p
        Call GenGraph(i, 30, 30, 600, 400)
        '���̃V�[�g��
        i = i + 1
    Loop
End Sub

Sub ReadCSVFile(FolderName As String, FileName As String)
    Dim buf As String
    Dim FullFilePath As String
    Dim tmp As Variant
    
    FullFilePath = FolderName & FileName
    With ActiveSheet
        .Select
        Open FullFilePath For Input As #1
            i = 0
            Do Until EOF(1)
                i = i + 1
                Line Input #1, buf
                tmp = Split(buf, ",")
                '�e�Z�����Ɠ���
                Dim j As Integer
                For j = 1 To UBound(tmp) + 1
                    .Cells(i, j).Value = tmp(j - 1)
                    '�ϊ�(String=>Integer)
                    Val (.Cells(i, j))
                Next j
            Loop
        Close #1
    End With
End Sub
    

