VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUpdateSettings 
   Caption         =   "OpenSolver - Update Settings"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "frmUpdateSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmUpdateSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If Mac Then
    Const FormWidthUpdateSettings = 350
#Else
    Const FormWidthUpdateSettings = 255
#End If

Private Sub chkEnabled_Change()
    chkExperimental.Enabled = chkEnabled.value
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOk_Click()
    SaveUpdateSetting chkEnabled.value
    SaveBetaUpdateSetting chkExperimental.value
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    lblUserAgent.Caption = GetUserAgent()
    chkEnabled.value = GetUpdateSetting()
    chkExperimental.value = GetBetaUpdateSetting()
    chkEnabled_Change
    
    AutoLayout
End Sub

Private Sub AutoLayout()
    AutoFormat Me.Controls
    
    Me.width = FormWidthUpdateSettings
    
    With lblDesc
        .Caption = "OpenSolver can automatically check for updates and let you know when a new version is available. " & _
                   "We only check for updates if OpenSolver is being actively used, and not more than once a day."
        .left = FormMargin
        .top = FormMargin
        AutoHeight lblDesc, Me.width - 2 * FormMargin
    End With
    
    With chkEnabled
        .width = lblDesc.width
        .left = lblDesc.left
        .top = Below(lblDesc)
    End With
    
    With chkExperimental
        .width = lblDesc.width
        .left = lblDesc.left
        .top = Below(chkEnabled, False)
    End With
    
    With lblInfo
        .Caption = "Our update check sends anonymous version information which lets us collect statistics on the " & _
                   "Operating System, Excel, and OpenSolver versions being used. This helps ensure we are testing " & _
                   "OpenSolver on all popular platforms. If you enable update checks, the information sent for your installation would be:"
        .left = lblDesc.left
        .top = Below(chkExperimental)
        AutoHeight lblInfo, lblDesc.width
    End With
    
    With lblUserAgent
        .left = lblDesc.left
        .top = Below(lblInfo)
        AutoHeight lblUserAgent, lblDesc.width
        .TextAlign = fmTextAlignCenter
        .BackColor = FormBackColor
    End With
    
    With cmdCancel
        .Caption = "Cancel"
        .width = FormButtonWidth
        .left = LeftOfForm(Me.width, .width)
        .top = Below(lblUserAgent)
    End With
    
    With cmdOk
        .Caption = "OK"
        .width = FormButtonWidth
        .left = LeftOf(cmdCancel, .width)
        .top = cmdCancel.top
    End With
        
    Me.height = FormHeight(cmdOk)
    Me.width = Me.width + FormWindowMargin
    
    Me.BackColor = FormBackColor
    Me.Caption = "OpenSolver - Update Settings"
End Sub
