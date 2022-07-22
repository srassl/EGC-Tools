Attribute VB_Name = "Grid"
'©2016 Efficient Elements GmbH
'Contact us at info@efficient-elements.com

Option Explicit

Public Sub SnapToGrid()

  On Error GoTo errHandler
  
  Dim ps As PageSetup
  Dim Sld As Slide
  Dim shp As Shape
  
  Dim gridSize As Single
  
  With ActivePresentation
    Set ps = .PageSetup
    gridSize = .GridDistance
  End With
  
  With ActiveWindow.Selection
    If .Type = ppSelectionSlides Then
      For Each Sld In .SlideRange
        For Each shp In Sld.Shapes
          If shp.Visible Then
            SnapShapeToGrid shp, gridSize, ps
          End If
        Next
      Next
    ElseIf .Type = ppSelectionShapes Or .Type = ppSelectionText Then
      For Each shp In .ShapeRange
        SnapShapeToGrid shp, gridSize, ps
      Next
    End If
  End With
  
  Exit Sub
errHandler:

End Sub

Private Sub SnapShapeToGrid(shp As Shape, gridSize As Single, ps As PageSetup)

  Dim x As Single
  Dim y As Single

  'Size
  shp.Width = Round(shp.Width / gridSize, 0) * gridSize
  shp.Height = Round(shp.Height / gridSize, 0) * gridSize
  
  'Position - we have to translate the coordinate system as the grid origin is in the center of the slide
  x = shp.Left - ps.SlideWidth / 2
  y = shp.Top - ps.SlideHeight / 2
  
  x = Round(x / gridSize, 0) * gridSize
  x = x + ps.SlideWidth / 2
  
  y = Round(y / gridSize, 0) * gridSize
  y = y + ps.SlideHeight / 2
  
  shp.Left = x
  shp.Top = y
  
End Sub
