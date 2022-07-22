Attribute VB_Name = "Footer"
Option Explicit

Public Sub CheckAtLoad()
  
  On Error GoTo checkAtLoadNotDefined
    
  If ActivePresentation.CustomDocumentProperties.Item("AutoONOFF").Value = True _
    Or ActivePresentation.CustomDocumentProperties.Item("AutoONOFF").Value = False Then
      
      If ActivePresentation.BuiltInDocumentProperties.Item("Title") = "Folienmaster EUROGROUP CONSULTING" Then

checkAtLoadNotDefined:
        If (InStr(ActivePresentation.TemplateName, "EGC") > 0) Then
          FooterDialog.Show
        End If
      End If
  End If
    
End Sub

Public Sub UpdateFileInfo()
    On Error Resume Next
    Dim spe As Shape
    Dim found As Boolean
    found = False
    
    'Wenn Fußzeilenautomatik an ist
    If GetCustomDocumentProperty("AutoONOFF") Then
        
        'Folien-Master updaten
        For Each spe In ActivePresentation.SlideMaster.Shapes
            If Left(spe.Name, 9) = "Rectangle" And spe.Top > 490 And spe.Left > 500 And spe.Type <> msoPlaceholder Then
                UpdateFooterShape spe
                found = True
                Exit For
            ElseIf spe.Tags("EGCFuss") = "1" Then
                UpdateFooterShape spe
                found = True
                Exit For
            End If
        Next
        'Wenn kein EGCFuß-Objekt gefunden wurde
        If Not found Then
            Select Case MsgBox("Das EGC-Fußzeilen-Objekt wurde im Master nicht gefunden." & vbCr & _
                      "Soll ein EGC-Fuß erstellt werden?" & vbCr & vbCr & _
                      "Ja: Eine EGC-Fußzeile wird im Master erstellt." & vbCr & vbCr & _
                      "Nein: Es wird KEINE Fußzeile erstellt und die Fußzeilenautomatik ABGESCHALTET." & vbCr & vbCr & _
                      "Abbrechen: Es wird KEINE Fußzeile erstellt.", vbQuestion + vbYesNoCancel, "EGC-Fuß")
                Case vbYes
                    UpdateFooterShape createEGCFooter()
                    PPTFooterOff
                Case vbNo
                    ActivePresentation.CustomDocumentProperties("AutoONOFF") = False
                    
            End Select
        End If
    End If
     
End Sub

Function UpdateFooterShape(spe As Shape)
    Dim FusszeilenText As String
    Dim Custom As DocumentProperties
    Dim BuiltIn As DocumentProperties
    Set Custom = ActivePresentation.CustomDocumentProperties
    Set BuiltIn = ActivePresentation.BuiltInDocumentProperties
    Dim offset As Integer
        
    On Error GoTo ChangeShapeNew_err1
    
    If Not GetCustomDocumentProperty("AutoONOFF") Then
        Exit Function
    End If
    
    '********* TITEL
    FusszeilenText = BuiltIn("Title").Value & " / "
    '********* Version
    If Custom("VersionONOFF") Then FusszeilenText = FusszeilenText & "Version " & GetVersion & " / "
    
    '********* Author
    FusszeilenText = FusszeilenText & BuiltIn("Author") & vbCr
    '********* Kunde
    If Len(Custom("Customer").Value) > 0 Then FusszeilenText = FusszeilenText & Custom("Customer").Value & " / "
    
    '********* Projektname
    If Len(BuiltIn("Subject").Value) > 0 Then FusszeilenText = FusszeilenText & BuiltIn("Subject").Value & " / "
    
    '********* Projektnummer
    If Len(Custom("ProjectNr").Value) > 0 Then FusszeilenText = FusszeilenText & Custom("ProjectNr").Value & " / "
                
    
    '********* Sprachabhängige Fusszeilenelemente
    '********* Deutsch
    If Not Custom("Language") Then
        
        If Custom("StandONOFF") Then FusszeilenText = FusszeilenText & "Stand " & Left(GetDatee, 10) & " / "
                
        If Custom("SeitVonONOFF") Then
            FusszeilenText = FusszeilenText & "Seite  von " & Application.ActivePresentation.Slides.Count
        Else
            FusszeilenText = FusszeilenText & "Seite "
        End If
    '********* English
    Else
        If Custom("StandONOFF") Then
            FusszeilenText = FusszeilenText & "Date " & Mid(GetDatee, 4, 2) & "/" & Mid(GetDatee, 1, 2) & "/" & Mid(GetDatee, 7, 4) & " / "
        End If
        
        If Custom("SeitVonONOFF") Then
            FusszeilenText = FusszeilenText & "Page  of " & Application.ActivePresentation.Slides.Count
        Else
            FusszeilenText = FusszeilenText & "Page "
        End If
    
    End If
    
    'Fußzeiletext anwenden
    spe.TextFrame.TextRange.Text = FusszeilenText
    
    'Die Seitennummer wird nachträglich eingefügt
    If Custom("SeitVonONOFF") Then
        If (Application.ActivePresentation.Slides.Count < 10) Then
            offset = 5
        ElseIf (Application.ActivePresentation.Slides.Count < 100) Then
            offset = 6
        Else
            offset = 7
        End If
        
        If Custom("Language") Then
            offset = offset - 1
        End If
        spe.TextFrame.TextRange.Characters(Len(FusszeilenText) - offset, 0).InsertSlideNumber
    Else
        spe.TextFrame.TextRange.Characters(Len(FusszeilenText) + 1, 0).InsertSlideNumber
        
    End If
    
ChangeShapeNew_err1:
    
End Function

Function createEGCFooter() As Shape
    On Error GoTo createEGCFooter_err
 
    deleteEGCFooter
    
    Set createEGCFooter = ActivePresentation.SlideMaster.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 100, 10, 100)
    
    createEGCFooter.Name = "EGCFuss"
    createEGCFooter.Tags.Add "EGCFuss", 1
    
    With createEGCFooter
        .Width = 351.5
        .Height = 31.75
        .Left = 38.5
        .Top = 508.25
        .TextFrame.TextRange.Text = "[EGCFuss]"
        .TextFrame.TextRange.Font.Name = "Arial"
        .TextFrame.TextRange.Font.Size = 9
        .TextFrame.TextRange.Font.Color.RGB = EGC_ANTHRAZIT
        .Fill.Visible = False
        .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
        .TextFrame.MarginBottom = 0
        .TextFrame.MarginLeft = 0
        .TextFrame.MarginRight = 0
        .TextFrame.MarginTop = 0
    End With
    
    Exit Function
createEGCFooter_err:
    createEGCFooter = Nothing
End Function

Function deleteEGCFooter()
    On Error Resume Next
    Dim spe As Shape
    
    For Each spe In ActivePresentation.SlideMaster.Shapes
        If spe.Name = "EGCFuss" Or spe.Tags("EGCFuss") = "1" Then spe.Delete
    Next
    
End Function

Function GetDatee()
    GetDatee = Application.ActivePresentation.BuiltInDocumentProperties("Last Save Time")
End Function

Function PPTFooterOff()
  Dim slideObj As Slide
  Dim customLayoutObj As CustomLayout
  Dim shapeObj As Shape
  
  On Error Resume Next
  'Powerpoint-Fusszeile auf allen Slides ausschalten
  For Each slideObj In ActivePresentation.Slides
    slideObj.HeadersFooters.Footer.Visible = False
    slideObj.HeadersFooters.SlideNumber.Visible = False
    slideObj.HeadersFooters.DateAndTime.Visible = False
  Next
  
      
  'PPT-Fußzeilen-Objekt im SlideMaster löschen
  For Each shapeObj In ActivePresentation.SlideMaster.Shapes
      If shapeObj.Type = msoPlaceholder Then
          If shapeObj.PlaceholderFormat.Type = ppPlaceholderFooter Then
              shapeObj.Delete
          End If
      End If
  Next
  
  ''PPT-Fußzeilen-Objekt in den Layouts löschen
  For Each customLayoutObj In ActivePresentation.SlideMaster.CustomLayouts
       For Each shapeObj In customLayoutObj.Shapes
          If shapeObj.Type = msoPlaceholder Then
              If shapeObj.PlaceholderFormat.Type = ppPlaceholderFooter Then
                  shapeObj.Delete
              End If
          End If
       Next
  Next
    
End Function

Private Function GetVersion() As String

  Dim regEx As New RegExp
  Dim matches As Object
  
  With regEx
    .Global = True
    .MultiLine = True
    .IgnoreCase = True
    .Pattern = "v[ \.](\d{1,2}\.\d{1,2})\.pptx$"
    Set matches = .Execute(ActivePresentation.Name)
    
    If matches.Count > 0 Then
      GetVersion = matches.Item(0).SubMatches.Item(0)
    End If
  End With
  
  Set matches = Nothing
  Set regEx = Nothing
  
End Function
