Attribute VB_Name = "ChartColoring"
' VBA Macros for Excel
' Functions to label charts
'
' Autor: Thomas Arend
' (c) 2018
' Stand: 16.01.2018

' Change then fill of the SeriesCollection 1 and 2 in a chart to light green and red

Private Sub ChngColorSeries(myChart As Chart)
 
If myChart.ChartType = -4111 And myChart.SeriesCollection.Count = 3 Then
  myChart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(128, 255, 128)
  myChart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 128, 128)
End If

End Sub

Sub ActiveChartChngColerSeries()

If Not ActiveChart Is Nothing Then
  
   Call ChngColorSeries(ActiveChart)
  
Else

  MsgBox "Bitte ein Diagramm auswählen!", vbOKOnly, "Fehler: Kein aktives Diagramm"
   
End If

End Sub
