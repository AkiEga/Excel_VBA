Attribute VB_Name = "etc"
Option Explicit

Sub 個人情報の削除可否を設定する()              '※１
    ActiveWorkbook.RemovePersonalInformation = True  '削除可能にする ※２
    ActiveWorkbook.RemovePersonalInformation = False '削除不可能にする
    MsgBox ActiveWorkbook.RemovePersonalInformation  '削除可否を取得する
End Sub

Sub SheetClear()
    Dim i As Variant
    For Each i In ThisWorkbook.Sheets
        
        If i.Name <> "Sheet1" Then
            Application.DisplayAlerts = False
            i.Delete
            Application.DisplayAlerts = True
        End If
        
    Next i
End Sub
