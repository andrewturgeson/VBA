Sub PMEFluidityPlots()
'Dim ArrayA As Variant
'ArrayA = Array("P", "L", "B", "Y", "H", "E", "A", "D")

'For i = 0 To 7
    'Set ws = ActiveWorkbook.Sheets(ArrayA(i) & "-All")
    Set ws = ActiveWorkbook.Sheets("P-All")
    Set rng = ActiveSheet.Range("B1:Q1500")
    Set cht = ActiveSheet.ChartObjects.Add( _
    Left:=ActiveCell.Left, _
    Width:=450, _
    Top:=ActiveCell.Top, _
    Height:=250)

'Give chart some data
  cht.Chart.SetSourceData Source:=rng

'Determine the chart type
  cht.Chart.ChartType = xlXYScatterLinesNoMarkers
'Next i

End Sub
