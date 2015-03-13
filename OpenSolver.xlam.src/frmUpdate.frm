VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUpdate 
   Caption         =   "OpenSolver - Update Available"
   ClientHeight    =   3164
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4725
   OleObjectBlob   =   "frmUpdate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If Mac Then
    Const FormWidthUpdate = 300
#Else
    Const FormWidthUpdate = 195
#End If

Private Sub chkKeepChecking_Change()
    SaveUpdateSetting chkKeepChecking.value
End Sub

Private Sub cmdOk_Click()
    Me.Hide
End Sub

Private Sub lblLink_Click()
    OpenURL lblLink.Caption
End Sub

Private Sub UserForm_Activate()
   chkKeepChecking.value = GetUpdateSetting()
End Sub

Private Sub UserForm_Initialize()
   AutoLayout
End Sub

Private Sub AutoLayout()
    AutoFormat Me.Controls
    
    Me.width = FormWidthUpdate
    
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
    
    With chkKeepChecking
        .Caption = "Continue checking for updates to OpenSolver"
        .width = lblDesc.width
        .top = Below(lblLink, False)
        .left = lblDesc.left
    End With
    
    With cmdOk
        .Caption = "OK"
        .width = FormButtonWidth
        .left = (Me.width - .width) / 2
        .top = Below(chkKeepChecking)
    End With
        
    Me.height = FormHeight(cmdOk)
    Me.width = Me.width + FormWindowMargin
    
    Me.BackColor = FormBackColor
    Me.Caption = "OpenSolver - Update Available"
End Sub

Sub ShowUpdate(LatestVersion As String)
    lblLatestVersion.Caption = "Latest version: " & LatestVersion
    lblCurrentVersion.Caption = "Current version: " & sOpenSolverVersion
    Me.Show
End Sub
