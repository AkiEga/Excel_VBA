Attribute VB_Name = "Module1"
'参考サイト: http://oshiete.goo.ne.jp/qa/3168255.html?from=recommend
Sub CSVまとめsample()
    Dim MyObj As Object
    Dim MyFol As String
    Dim MyFnm As String
    Dim MyStr As String
    Dim i     As Long
    Dim n     As Long
    Dim n1    As Long
  
    'フォルダを選択する
    'Set MyObj = CreateObject("Shell.Application") _
    '    .BrowseForFolder(0, "SelectFolder", 0, ActiveWorkbook.Path)
    'Set MyObj = CreateObject("Shell.Application") _
    '選択なければ処理を抜ける
    'If MyObj Is Nothing Then Exit Sub
    
    MyFol = Worksheets("Sheet1").Range("D2").Value
    'フォルダが存在しないならば再度フォルダ選択
    If Dir(MyFol) = "" Then
        Module3.SelectDir
        MyFol = Worksheets("Sheet1").Range("D2").Value
    End If
    
    MsgBox MyFol & "を処理します。"
    
    If ActiveWorkbook.Sheets.Count >= 2 Then
        Module2.SheetClear
    End If
    'オートフィルターの設定値
    Dim AutoFilMin As String
    Dim AutoFilMax As String
    AutoFilMin = ">=" + CStr(Worksheets("Sheet1").Range("B3").Value)
    AutoFilMax = "<=" + CStr(Worksheets("Sheet1").Range("B4").Value)
    MsgBox ("AutoFilMin=" + AutoFilMin + ",AutoFilMax=" + AutoFilMax)
    i = 0
    'Dir関数を使って指定フォルダ内csvファイルを順次処理
    MyFnm = Dir(MyFol & "*.csv")
    Do Until Len(MyFnm) = 0&
        '何個目のファイルを取り込んだかカウント
        i = i + 1
        'ThisWorkbookにシートを追加して処理
        With Sheets.Add(after:=Worksheets(Worksheets.Count))
            'ファイル名からシート名を抽出
            .Name = Mid(MyFnm, 8, 12)
            '外部データ取り込みを利用
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
            
           
            'オートフィルターを追加
            .Range("A1").AutoFilter _
            Field:=3, _
            Criteria1:=AutoFilMin, Operator:=xlAnd, Criteria2:=AutoFilMax
            .Range("A1").AutoFilter _
            Field:=4, _
            Criteria1:=AutoFilMin, Operator:=xlAnd, Criteria2:=AutoFilMax
        
        End With
        '次のファイルへ
        MyFnm = Dir()
    Loop
    If i > 0 Then
        MyStr = i & "個のファイルを処理しました。"
    Else
        '検索結果が０なら
        MyStr = "検索条件を満たすファイルはありません。"
    End If
    Application.ScreenUpdating = True
    MsgBox MyStr
End Sub

