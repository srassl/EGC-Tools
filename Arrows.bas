Attribute VB_Name = "Arrows"
'©2016 Efficient Elements GmbH
'Contact us at info@efficient-elements.com

Option Explicit

Type lineType
  beginX As Single
  beginY As Single
  endX As Single
  endY As Single
End Type

Sub AddArrow(arrowType As Integer)

  On Error GoTo errHandler

  Dim gridSize As Single

  Dim shp As ShapeRange
  Dim Sld As Slide
  
  Dim triangleArray(1 To 4, 1 To 2) As Single
  Dim line1 As lineType
  Dim line2 As lineType
  
  gridSize = ActivePresentation.GridDistance
  Set shp = ActiveWindow.Selection.ShapeRange
  Set Sld = ActiveWindow.View.Slide
  
  With shp
    Select Case arrowType
      'Top
      Case 0
        triangleArray(1, 1) = .Left + .Width / 2 - 3 * gridSize
        triangleArray(1, 2) = .Top - gridSize
        triangleArray(2, 1) = .Left + .Width / 2
        triangleArray(2, 2) = .Top - 3 * gridSize
        triangleArray(3, 1) = .Left + .Width / 2 + 3 * gridSize
        triangleArray(3, 2) = .Top - gridSize
        triangleArray(4, 1) = .Left + .Width / 2 - 3 * gridSize
        triangleArray(4, 2) = .Top - gridSize
        line1.beginX = .Left
        line1.beginY = .Top - 2 * gridSize
        line1.endX = .Left + .Width / 2 - 4 * gridSize
        line1.endY = line1.beginY
        line2.beginX = .Left + .Width / 2 + 4 * gridSize
        line2.beginY = line1.beginY
        line2.endX = .Left + .Width
        line2.endY = line1.beginY
      
      'Right
      Case 1
        triangleArray(1, 1) = .Left + .Width + gridSize
        triangleArray(1, 2) = .Top + .Height / 2 - 3 * gridSize
        triangleArray(2, 1) = .Left + .Width + 3 * gridSize
        triangleArray(2, 2) = .Top + .Height / 2
        triangleArray(3, 1) = .Left + .Width + gridSize
        triangleArray(3, 2) = .Top + .Height / 2 + 3 * gridSize
        triangleArray(4, 1) = .Left + .Width + gridSize
        triangleArray(4, 2) = .Top + .Height / 2 - 3 * gridSize
        line1.beginX = .Left + .Width + 2 * gridSize
        line1.beginY = .Top
        line1.endX = line1.beginX
        line1.endY = .Top + .Height / 2 - 4 * gridSize
        line2.beginX = line1.beginX
        line2.beginY = .Top + .Height / 2 + 4 * gridSize
        line2.endX = line1.beginX
        line2.endY = .Top + .Height
      
      'Bottom
      Case 2
        triangleArray(1, 1) = .Left + .Width / 2 - 3 * gridSize
        triangleArray(1, 2) = .Top + .Height + gridSize
        triangleArray(2, 1) = .Left + .Width / 2
        triangleArray(2, 2) = .Top + .Height + 3 * gridSize
        triangleArray(3, 1) = .Left + .Width / 2 + 3 * gridSize
        triangleArray(3, 2) = .Top + .Height + gridSize
        triangleArray(4, 1) = .Left + .Width / 2 - 3 * gridSize
        triangleArray(4, 2) = .Top + .Height + gridSize
        line1.beginX = .Left
        line1.beginY = .Top + .Height + 2 * gridSize
        line1.endX = .Left + .Width / 2 - 4 * gridSize
        line1.endY = line1.beginY
        line2.beginX = .Left + .Width / 2 + 4 * gridSize
        line2.beginY = line1.beginY
        line2.endX = .Left + .Width
        line2.endY = line1.beginY
      
      'Left
      Case 3
      triangleArray(1, 1) = .Left - gridSize
      triangleArray(1, 2) = .Top + .Height / 2 - 3 * gridSize
      triangleArray(2, 1) = .Left - 3 * gridSize
      triangleArray(2, 2) = .Top + .Height / 2
      triangleArray(3, 1) = .Left - gridSize
      triangleArray(3, 2) = .Top + .Height / 2 + 3 * gridSize
      triangleArray(4, 1) = .Left - gridSize
      triangleArray(4, 2) = .Top + .Height / 2 - 3 * gridSize
      line1.beginX = .Left - 2 * gridSize
      line1.beginY = .Top
      line1.endX = line1.beginX
      line1.endY = .Top + .Height / 2 - 4 * gridSize
      line2.beginX = line1.beginX
      line2.beginY = .Top + .Height / 2 + 4 * gridSize
      line2.endX = line1.beginX
      line2.endY = .Top + .Height
    End Select
  End With
  
  With Sld.Shapes
    .AddLine(line1.beginX, line1.beginY, line1.endX, line1.endY).Select True
    .AddLine(line2.beginX, line2.beginY, line2.endX, line2.endY).Select False
  End With
  
  With ActiveWindow.Selection.ShapeRange.Line
    .Visible = msoTrue
    .ForeColor.RGB = EGC_RED
    .Weight = 1#
    .DashStyle = msoLineSolid
  End With
  
  With Sld.Shapes.AddPolyline(triangleArray)
    .Select False
    .Line.Visible = False
    .Fill.Visible = msoTrue
    .Fill.ForeColor.RGB = EGC_RED
    .Fill.Solid
  End With
  
  ActiveWindow.Selection.ShapeRange.Group.Select msoTrue
  shp.Select True
    
  Exit Sub
errHandler:

End Sub
