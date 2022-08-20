Sub addChartsSSOC()
'Keyboard Shortcut: Ctrl+q

Dim rng As Range

Dim Ax As Integer
Dim Ay As Integer
Dim Az As Integer
Ax = ActiveCell.Column
Ay = ActiveCell.Row
Az = Cells(Rows.Count, ActiveCell.Column).End(xlUp).Row
Set rng = ActiveSheet.Range(Cells(Ay, Ax), Cells(Az, Ax + 3))
  
ActiveSheet.Shapes.AddChart2(240, xlXYScatterLinesNoMarkers).Select
    ActiveChart.SetSourceData Source:=rng
    With ActiveChart.Parent
        .Left = ActiveCell.Left + 20
        .Top = ActiveCell.Top + 20
    End With

End Sub
