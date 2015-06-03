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
    
    Me.width = FormWidthUpdateNotification
    
    With lblDesc
        .Caption = "A newer version of OpenSolver is available. Please follow the link below for more information and to download the update:"
        .left = FormMargin
        .top = FormMargin
        AutoHeight lblDesc, Me.width - 2 * FormMargin
    End With
    
    With lblLatestVersion
        .width = lblDesc.width / 2
        .left = lblDesc.left
        .top = Below(lblDesc)
    End With
    
    With lblCurrentVersion
        .width = lblLatestVersion.width
        .left = RightOf(lblLatestVersion, False)
        .top = lblLatestVersion.top
    End With
    
    With lblLink
        .Caption = "http://OpenSolver.org"
        .ForeColor = FormLinkColor
        .Font.Underline = True
        .left = lblDesc.left
        .top = Below(lblLatestVersion, False)
        .width = lblDesc.width
        .TextAlign = fmTextAlignCenter
    End With
    
    With cmdSettings
        .Caption = "Update Settings..."
        .width = (lblDesc.width - FormSpacing) / 2
        .left = lblDesc.left
        .top = Below(lblLink)
    End With
    
    With cmdOk
        .Caption = "OK"
        .left = RightOf(cmdSettings)
        .width = cmdSettings.width
        .top = cmdSettings.top
    End With
        
    Me.height = FormHeight(cmdOk)
    Me.width = Me.width + FormWindowMargin
    
    Me.BackColor = FormBackColor
    Me.Caption = "OpenSolver - Update Available"
End Sub

Private Sub CenterForm()
    Me.top = CenterFormTop(Me.height)
    Me.left = CenterFormLeft(Me.width)
End Sub

Sub ShowUpdate(LatestVersion As String)
    lblLatestVersion.Caption = "Latest version: " & LatestVersion
    lblCurrentVersion.Caption = "Current version: " & sOpenSolverVersion
    Me.Show
End Sub
