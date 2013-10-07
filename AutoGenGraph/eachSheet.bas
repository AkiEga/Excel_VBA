Attribute VB_Name = "eachSheet"
'参考サイト: http://oshiete.goo.ne.jp/qa/3168255.html?from=recommend
Dim MyObj As Object
Dim MyFol As String
Dim MyFnm As String
Dim MyStr As String
Dim i   As Integer
Dim n   As Long
Dim n1  As Long

Sub EachSheetReadCSV()
    'もうすでにシートを作っていたら再度作り直し
    If ActiveWorkbook.Sheets.Count >= 2 Then
        SheetClear
    End If
    'フォルダを選択する
    MyFol = SetCSVFolder
    MsgBox MyFol & "を処理します。"
    
    'Dir関数を使って指定フォルダ内csvファイルを順次処理
    MyFnm = Dir(MyFol & "*.csv")
    i = 0
    Do Until Len(MyFnm) = 0&
        '何個目のファイルを取り込んだかカウント
        i = i + 1
        'ThisWorkbookにシートを追加して処理
        With Sheets.Add(after:=Worksheets(Worksheets.Count))
            'ファイル名からシート名を抽出
            .Name = Mid(MyFnm, 1, Len(MyFnm) - Len(".csv"))
            '外部データ取り込みを利用
            Call ReadCSVFile(MyFol, MyFnm)
        End With
        '次のファイルへ
        MyFnm = Dir()
    Loop

End Sub

Sub EachSheetGenGraph()
    i = 2
    Do Until i > Sheets.Count
        Debug.Print CStr(i) + "シート目のグラフを作成します"
        '何個目のファイルを取り込んだかカウント
        '外部データ取り込みを利用
        Call GenGraph(i, 30, 30, 600, 400)
        '次のシートへ
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
                '各セルごと入力
                Dim j As Integer
                For j = 1 To UBound(tmp) + 1
                    .Cells(i, j).Value = tmp(j - 1)
                    '変換(String=>Integer)
                    Val (.Cells(i, j))
                Next j
            Loop
        Close #1
    End With
End Sub
    

