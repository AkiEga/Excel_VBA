VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub CommandButton1_Click()
    Call EachSheetGenGraph
End Sub

Private Sub CommandButton2_Click() '※１
    ActiveWorkbook.RemovePersonalInformation = True  '削除可能にする ※２
    'ActiveWorkbook.RemovePersonalInformation = False '削除不可能にする
    MsgBox ActiveWorkbook.RemovePersonalInformation  '削除可否を取得する
End Sub

Private Sub CommandButton3_Click()
    Call SetCSVFolder
End Sub

Private Sub CommandButton4_Click()
    Call EachSheetReadCSV
End Sub
