Attribute VB_Name = "Module1"
'�Q�l�T�C�g: http://oshiete.goo.ne.jp/qa/3168255.html?from=recommend
Sub CSV�܂Ƃ�sample()
    Dim MyObj As Object
    Dim MyFol As String
    Dim MyFnm As String
    Dim MyStr As String
    Dim i     As Long
    Dim n     As Long
    Dim n1    As Long
  
    '�t�H���_��I������
    'Set MyObj = CreateObject("Shell.Application") _
    '    .BrowseForFolder(0, "SelectFolder", 0, ActiveWorkbook.Path)
    'Set MyObj = CreateObject("Shell.Application") _
    '�I���Ȃ���Ώ����𔲂���
    'If MyObj Is Nothing Then Exit Sub
    
    MyFol = Worksheets("Sheet1").Range("D2").Value
    '�t�H���_�����݂��Ȃ��Ȃ�΍ēx�t�H���_�I��
    If Dir(MyFol) = "" Then
        Module3.SelectDir
        MyFol = Worksheets("Sheet1").Range("D2").Value
    End If
    
    MsgBox MyFol & "���������܂��B"
    
    If ActiveWorkbook.Sheets.Count >= 2 Then
        Module2.SheetClear
    End If
    '�I�[�g�t�B���^�[�̐ݒ�l
    Dim AutoFilMin As String
    Dim AutoFilMax As String
    AutoFilMin = ">=" + CStr(Worksheets("Sheet1").Range("B3").Value)
    AutoFilMax = "<=" + CStr(Worksheets("Sheet1").Range("B4").Value)
    MsgBox ("AutoFilMin=" + AutoFilMin + ",AutoFilMax=" + AutoFilMax)
    i = 0
    'Dir�֐����g���Ďw��t�H���_��csv�t�@�C������������
    MyFnm = Dir(MyFol & "*.csv")
    Do Until Len(MyFnm) = 0&
        '���ڂ̃t�@�C������荞�񂾂��J�E���g
        i = i + 1
        'ThisWorkbook�ɃV�[�g��ǉ����ď���
        With Sheets.Add(after:=Worksheets(Worksheets.Count))
            '�t�@�C��������V�[�g���𒊏o
            .Name = Mid(MyFnm, 8, 12)
            '�O���f�[�^��荞�݂𗘗p
            With .QueryTables.Add(Connection:="TEXT;" & MyFol & MyFnm, _
                                  Destination:=.Range("A" & 1))
                .AdjustColumnWidth = False
                .TextFilePlatform = xlWindows
                .TextFileStartRow = 1
                .TextFileCommaDelimiter = True
                .Refresh False
                n1 = .ResultRange.Rows.Count
                .Parent.Names(.Name).Delete
                .Delete
            End With
            
           
            '�I�[�g�t�B���^�[��ǉ�
            .Range("A1").AutoFilter _
            Field:=3, _
            Criteria1:=AutoFilMin, Operator:=xlAnd, Criteria2:=AutoFilMax
            .Range("A1").AutoFilter _
            Field:=4, _
            Criteria1:=AutoFilMin, Operator:=xlAnd, Criteria2:=AutoFilMax
        
        End With
        '���̃t�@�C����
        MyFnm = Dir()
    Loop
    If i > 0 Then
        MyStr = i & "�̃t�@�C�����������܂����B"
    Else
        '�������ʂ��O�Ȃ�
        MyStr = "���������𖞂����t�@�C���͂���܂���B"
    End If
    Application.ScreenUpdating = True
    MsgBox MyStr
End Sub

