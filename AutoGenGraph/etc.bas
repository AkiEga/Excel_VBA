Attribute VB_Name = "etc"
Option Explicit

Sub �l���̍폜�ۂ�ݒ肷��()              '���P
    ActiveWorkbook.RemovePersonalInformation = True  '�폜�\�ɂ��� ���Q
    ActiveWorkbook.RemovePersonalInformation = False '�폜�s�\�ɂ���
    MsgBox ActiveWorkbook.RemovePersonalInformation  '�폜�ۂ��擾����
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
