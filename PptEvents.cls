VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PptEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'©2016 Efficient Elements GmbH
'Contact us at info@efficient-elements.com

Public WithEvents PptApp As PowerPoint.Application
Attribute PptApp.VB_VarHelpID = -1

Private Sub PptApp_PresentationNewSlide(ByVal Sld As Slide)
    UpdateFileInfo
End Sub

Private Sub PptApp_PresentationOpen(ByVal Pres As Presentation)
  CheckAtLoad
End Sub

Private Sub PptApp_PresentationPrint(ByVal Pres As Presentation)
    UpdateFileInfo
End Sub

Private Sub PptApp_SlideShowBegin(ByVal Wn As SlideShowWindow)
    UpdateFileInfo
End Sub

'**Nicht Aktiviert wegen Automatisch Speichern auf SharePoint**
'Private Sub PptApp_PresentationSave(ByVal Pres As Presentation)
'  UpdateFileInfo
'End Sub


