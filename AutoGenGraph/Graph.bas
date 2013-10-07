Attribute VB_Name = "Graph"
Option Explicit
' Macro1 Macro
Sub GenGraph(SheetNum As Integer, X As Integer, Y As Integer, W As Integer, H As Integer)
Attribute GenGraph.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim oChart As ChartObject
    '�V�[�g��I��
    Worksheets(SheetNum).Select
    Worksheets(SheetNum).Activate
    '�O���t���������łɂ���ƏI��
    If ActiveSheet.ChartObjects.Count = 1 Then
        Exit Sub
    End If
    
    Set oChart = ActiveSheet.ChartObjects.Add(X, Y, W, H)
    
    With ActiveSheet.ChartObjects.Add(X, Y, W, H).Chart
        .ChartType = xlLineMarkers '�O���t�̎�ނ͐��O���t
        .SetSourceData Source:=GraphDataRange(SheetNum)  '�f�[�^�͈�
        .HasLegend = False '�}��͔�\��
        .FullSeriesCollection(1).Select
        .PlotArea.Select
        With ActiveChart.FullSeriesCollection(1)
            .Select
            .ApplyDataLabels
            .DataLabels.Select
             '�O���t�̒l���g��
            .DataLabels.Font.Size = 36
        End With
    End With

End Sub


Function GraphDataRange(SheetNum As Integer) As Range
    Dim sCell As Range
    Dim eCell As Range
    
    Set sCell = _
     Worksheets(SheetNum).Cells(Worksheets(1).Cells(3, 2).Value, _
                                Worksheets(1).Cells(3, 3).Value)
    Set eCell = _
     Worksheets(SheetNum).Cells(Worksheets(1).Cells(4, 2).Value, _
                                Worksheets(1).Cells(4, 3).Value)
    Set GraphDataRange = Worksheets(SheetNum).Range(sCell, eCell)
    
End Function
