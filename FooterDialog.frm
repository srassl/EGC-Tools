VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FooterDialog 
   Caption         =   "Einstellungen Fußzeile"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550.001
   OleObjectBlob   =   "FooterDialog.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "FooterDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Synopse Properties und Controls der Dialogbox
'        ActivePresentation.CustomDocumentProperties.Item("Customer").Value = FooterDialog.txtCustomer.Value
'        .Item("Subject").Value = FooterDialog.txtProject
'        ActivePresentation.CustomDocumentProperties.Item("ProjectNr").Value = FooterDialog.txtProjectNr
'        .Item("Title").Value = FooterDialog.txtTitle
'        .Item("Author").Value = FooterDialog.txtAuthor
'        ActivePresentation.CustomDocumentProperties.Item("VersionONOFF").Value = FooterDialog.chkVersion.Value
'        ActivePresentation.CustomDocumentProperties.Item("StandONOFF").Value = FooterDialog.chkStatus.Value
'        ActivePresentation.CustomDocumentProperties.Item("AutoONOFF").Value = FooterDialog.chkAuto.Value
'        ActivePresentation.CustomDocumentProperties.Item("SeitVonONOFF").Value = FooterDialog.chkSeiteVon.Value

Private Sub butTitle_Click()
    FooterDialog.txtTitle.SetFocus
End Sub

Private Sub chkAuto_Click()
    Me.setOptEnable
End Sub

Private Sub cmdcreateFootline_Click()
    UpdateFooterShape createEGCFooter
    forceViewUpdate
End Sub

Private Sub cmdDelFootline_Click()
    deleteEGCFooter
    forceViewUpdate
End Sub

Private Sub cmdOK_Click()

    Dim cont As TextBox

    With ActivePresentation.BuiltInDocumentProperties
        .Item("Subject").Value = FooterDialog.txtProject
        .Item("Title").Value = FooterDialog.txtTitle
        .Item("Author").Value = FooterDialog.txtAuthor
    End With
    
    With ActivePresentation.CustomDocumentProperties
        .Item("ProjectNr").Value = FooterDialog.txtProjectNr
        .Item("VersionONOFF").Value = FooterDialog.chkVersion
        .Item("StandONOFF").Value = FooterDialog.chkStatus
        .Item("AutoONOFF").Value = FooterDialog.chkAuto
        .Item("SeitVonONOFF").Value = FooterDialog.chkSeiteVon
        .Item("Customer").Value = FooterDialog.txtCustomer
        .Item("Language").Value = FooterDialog.chkLang
    End With
    
    Call UpdateFileInfo
    FooterDialog.Hide
    forceViewUpdate
End Sub

Private Sub cmdAbbrechen_Click()
    FooterDialog.Hide
End Sub



Private Sub UserForm_Activate()
    On Error Resume Next
    With ActivePresentation.BuiltInDocumentProperties
        FooterDialog.txtProject = .Item("Subject").Value
        FooterDialog.txtTitle = .Item("Title").Value
        FooterDialog.txtAuthor = .Item("Author").Value
    End With
    
    With ActivePresentation.CustomDocumentProperties
        .Add "ProjectNr", False, msoPropertyTypeString, "000000"
        .Add "Customer", False, msoPropertyTypeString, "NN"
        .Add "VersionONOFF", False, msoPropertyTypeBoolean, True
        .Add "StandONOFF", False, msoPropertyTypeBoolean, True
        .Add "AutoONOFF", False, msoPropertyTypeBoolean, True
        .Add "SeitVonONOFF", False, msoPropertyTypeBoolean, True
        .Add "Language", False, msoPropertyTypeBoolean, True
        
        FooterDialog.txtProjectNr = .Item("ProjectNr").Value
        FooterDialog.chkVersion = .Item("VersionONOFF").Value
        FooterDialog.chkStatus = .Item("StandONOFF").Value
        FooterDialog.chkAuto = .Item("AutoONOFF").Value
        FooterDialog.chkSeiteVon = .Item("SeitVonONOFF").Value
        FooterDialog.txtCustomer = .Item("Customer").Value
        FooterDialog.chkLang = .Item("Language").Value
    End With
    setOptEnable
End Sub

Public Function forceViewUpdate()
    Dim slideNum As Integer
    Dim slideCnt As Integer
    On Error GoTo Err_forceViewUpdate
    slideNum = ActiveWindow.View.Slide.SlideNumber
    slideCnt = ActivePresentation.Slides.Count
    If slideCnt > 1 Then
        If slideNum < slideCnt Then
            ActiveWindow.View.GotoSlide slideNum + 1
            ActiveWindow.View.GotoSlide slideNum
        Else
            ActiveWindow.View.GotoSlide slideNum - 1
            ActiveWindow.View.GotoSlide slideNum
        End If
    End If
Err_forceViewUpdate:
End Function
Function setOptEnable()
    
       Me.chkVersion.Enabled = chkAuto.Value
       Me.chkStatus.Enabled = chkAuto
       Me.chkSeiteVon.Enabled = chkAuto
       Me.chkLang.Enabled = chkAuto
       
End Function
