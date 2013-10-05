Attribute VB_Name = "Module1"
Option Explicit
' Macro1 Macro
Sub GenGraph(SheetNum As Integer, X As Integer, Y As Integer, W As Integer, H As Integer)
Attribute GenGraph.VB_ProcData.VB_Invoke_Func = " \n14"
    'シートを選択
    Worksheets(SheetNum).Select
    Worksheets(SheetNum).Activate
    'グラフがもうすでにあると終了
    If ActiveSheet.ChartObjects.Count = 1 Then
        Exit Sub
    End If
    
    With ActiveSheet.ChartObjects.Add(X, Y, W, H).Chart
        .ChartType = xlLineMarkers 'グラフの種類は線グラフ
        '.SetSourceData Source:=Worksheets(2).Range("E8:E44") 'データ範囲
        .SetSourceData Source:=GraphDataRange(SheetNum)  'データ範囲
        .HasLegend = False '凡例は非表示
        .FullSeriesCollection(1).Select
        .PlotArea.Select
        With ActiveChart.FullSeriesCollection(1)
            .Select
            .ApplyDataLabels
            .DataLabels.Select
             'グラフの値を拡大
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
