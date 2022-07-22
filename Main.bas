Attribute VB_Name = "Main"
'©2016 Efficient Elements GmbH
'Contact us at info@efficient-elements.com

Option Explicit

Private pptevts As PptEvents

Public Sub Auto_Open()

  Set pptevts = New PptEvents
  Set pptevts.PptApp = Application

End Sub

Public Sub OnAction(control As IRibbonControl)

  Select Case control.Id
    Case "ee_egc_1001"
      AddArrow 0
    Case "ee_egc_1002"
      AddArrow 2
    Case "ee_egc_1003"
      AddArrow 3
    Case "ee_egc_1004"
      AddArrow 1
    Case "ee_egc_1005"
      FooterDialog.Show
    Case "ee_egc_1006"
      SnapToGrid
  End Select

End Sub

Public Function GetCustomDocumentProperty(itemName As String) As Variant
  
  Dim prop As DocumentProperty
  
  For Each prop In ActivePresentation.CustomDocumentProperties
    If prop.Name = itemName Then
      GetCustomDocumentProperty = prop.Value
      Exit Function
    End If
  Next
  GetCustomDocumentProperty = ""
  
End Function
