Attribute VB_Name = "ChartLabeling"
' VBA Macros for Excel
' Functions to label charts
'
' Autor: Thomas Arend
' (c) 2018
' Stand: 16.01.2018

' Label a chart with a target-actual-value-comparison with %-value instead of the absolute data value.
'
' You need a chart type 52 with multile values and the sum of values in the last series
' The absolute values of the data labels are replace with the relative value of the sum (last series).
'

Private Sub PercentLabelChart(myChart As Chart)

Dim mySeries As Series
Dim TargetSeries As Series
Dim myPoint As Point

Dim i, j As Long

' Hide the sum series in the chart.

If myChart.ChartType = 52 Then
  
  With myChart.SeriesCollection(myChart.SeriesCollection.Count)
    .ChartType = xlLine
    .Format.Line.Visible = msoFalse
    .DataLabels.ShowValue = False
    
  End With
  
End If

'
' Now the chart must be Type -4111
' Set the labels to the %-values of the sum series
'

If myChart.ChartType = -4111 And myChart.SeriesCollection.Count > 2 Then
    
   Set TargetSeries = myChart.SeriesCollection(myChart.SeriesCollection.Count)
      
   For i = 1 To myChart.SeriesCollection.Count - 1
           
     Set mySeries = myChart.SeriesCollection(i)
        
     For j = 1 To mySeries.Points.Count
        
       Set myPoint = mySeries.Points(j)
        
       If myPoint.HasDataLabel Then
       ' myPoint.DataLabel.Text = Format(mySeries.Values(j), "0")
         If TargetSeries.Values(j) <> 0 Then
           myPoint.DataLabel.Text = Format(mySeries.Values(j) / TargetSeries.Values(j), "0%")
         End If
       Else
       ' If myPoint.Height > 120 Then
           myPoint.ApplyDataLabels Type:=xlDataLabelsShowValue
           myPoint.DataLabel.Text = Format(mySeries.Values(j) / TargetSeries.Values(j), "0%")
       ' End If
       End If
        
     Next ' j
      
   Next ' i
Else

  MsgBox ActiveChart.ChartType, vbOKOnly, "Wrong Chartyp: ChartType -4111 needed"
    
End If


End Sub

' Label all charts with a target-actual-value-comparison with %-value instead of the absolute value.

Sub ModifyAllCharts()

Dim MySheet As Worksheet

For Each MySheet In Worksheets

  If MySheet.ChartObjects.Count > 0 Then
    Call PercentLabelChart(MySheet.ChartObjects(1).Chart)
  End If
  
Next

End Sub

' Label the ActiceChart

Sub ActiveChartPercentLabeling()

If Not ActiveChart Is Nothing Then
  
   Call PercentLabelChart(ActiveChart)
  
Else

  MsgBox "Kein Diagramm ausgewählt. Bitte ein Diagramm auswählen!", vbOKOnly, "Fehler: Aktives Diagramm"
   
End If

End Sub


Private Sub ShowChartType()
  
If Not ActiveChart Is Nothing Then
  
   MsgBox ActiveChart.ChartType, vbOKOnly, "Diagrammtyp:"
     
Else

  MsgBox "Kein Diagramm ausgewählt. Bitte ein Diagramm auswählen!", vbOKOnly, "Fehler: Aktives Diagramm"
   
End If
  
End Sub

