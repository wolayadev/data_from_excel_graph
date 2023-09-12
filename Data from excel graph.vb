Sub GetChartValues()
   Dim NumRows As Integer
   Dim X As Object
   
   StartRow = 2
   SNum = 1

   Worksheets("ChartData").Cells(1, 1) = "Serie"
   Worksheets("ChartData").Cells(1, 2) = "X Values"
   Worksheets("ChartData").Cells(1, 3) = "Y Values"
   
   For Each X In ActiveChart.SeriesCollection
      
      NumRows = UBound(ActiveChart.SeriesCollection(SNum).Values)
            
      With Worksheets("ChartData")
         .Range(.Cells(StartRow, 1), .Cells(StartRow + NumRows - 1, 1)) = X.Name
      End With
      With Worksheets("ChartData")
         .Range(.Cells(StartRow, 2), .Cells(StartRow + NumRows - 1, 2)) = Application.Transpose(ActiveChart.SeriesCollection(SNum).XValues)
      End With
      With Worksheets("ChartData")
         .Range(.Cells(StartRow, 3), .Cells(StartRow + NumRows - 1, 3)) = Application.Transpose(X.Values)
      End With
      
      StartRow = StartRow + NumRows
      
      SNum = SNum + 1
   
   Next

End Sub
