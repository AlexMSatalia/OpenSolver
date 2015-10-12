VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FUpdateNotification 
   Caption         =   "OpenSolver - Update Available"
   ClientHeight    =   3164
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4725
   OleObjectBlob   =   "FUpdateNotification.frx":0000
End
Attribute VB_Name = "FUpdateNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If Mac Then
    Const FormWidthUpdateNotification = 300
#Else
    Const FormWidthUpdateNotification = 195
#End If

Private Sub cmdOk_Click()
    Me.Hide
End Sub

' Make the [x] hide the form rather than unload
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' If CloseMode = vbFormControlMenu then we know the user
    ' clicked the [x] close button or Alt+F4 to close the form.
    If CloseMode = vbFormControlMenu Then
        cmdOk_Click
        Cancel = True
    End If
End Sub

Private Sub cmdSettings_Click()
    Dim frmUpdateSettings As FUpdateSettings
    Set frmUpdateSettings = New FUpdateSettings
    frmUpdateSettings.Show
    Unload frmUpdateSettings
End Sub

Private Sub lblLink_Click()
    OpenURL lblLink.Caption
End Sub

Private Sub UserForm_Activate()
    CenterForm
End Sub

Private Sub UserForm_Initialize()
    AutoLayout
    CenterForm
End Sub

Private Sub AutoLayout()
    AutoFormat Me.Controls
    
    Me.Width = FormWidthUpdateNotification
    
    With lblDesc
        .Caption = "A newer version of OpenSolver is available. Please follow the link below for more information and to download the update:"
        .Left = FormMargin
        .Top = FormMargin
        AutoHeight lblDesc, Me.Width - 2 * FormMargin
    End With
    
    With lblLatestVersion
        .Width = lblDesc.Width / 2
        .Left = lblDesc.Left
        .Top = Below(lblDesc)
    End With
    
    With lblCurrentVersion
        .Width = lblLatestVersion.Width
        .Left = RightOf(lblLatestVersion, False)
        .Top = lblLatestVersion.Top
    End With
    
    With lblLink
        .Caption = "http://OpenSolver.org"
        .ForeColor = FormLinkColor
        .Font.Underline = True
        .Left = lblDesc.Left
        .Top = Below(lblLatestVersion, False)
        .Width = lblDesc.Width
        .TextAlign = fmTextAlignCenter
    End With
    
    With cmdSettings
        .Caption = "Update Settings..."
        .Width = (lblDesc.Width - FormSpacing) / 2
        .Left = lblDesc.Left
        .Top = Below(lblLink)
    End With
    
    With cmdOk
        .Caption = "OK"
        .Left = RightOf(cmdSettings)
        .Width = cmdSettings.Width
        .Top = cmdSettings.Top
        .Cancel = True
    End With
        
    Me.Height = FormHeight(cmdOk)
    Me.Width = Me.Width + FormWindowMargin
    
    Me.BackColor = FormBackColor
    Me.Caption = "OpenSolver - Update Available"
End Sub

Private Sub CenterForm()
    Me.Top = CenterFormTop(Me.Height)
    Me.Left = CenterFormLeft(Me.Width)
End Sub

Sub ShowUpdate(LatestVersion As String)
    lblLatestVersion.Caption = "Latest version: " & LatestVersion
    lblCurrentVersion.Caption = "Current version: " & sOpenSolverVersion
    Me.Show
End Sub
