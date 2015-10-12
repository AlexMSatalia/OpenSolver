VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FUpdateSettings 
   Caption         =   "OpenSolver - Update Settings"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "FUpdateSettings.frx":0000
End
Attribute VB_Name = "FUpdateSettings"
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

' Make the [x] hide the form rather than unload
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' If CloseMode = vbFormControlMenu then we know the user
    ' clicked the [x] close button or Alt+F4 to close the form.
    If CloseMode = vbFormControlMenu Then
        cmdCancel_Click
        Cancel = True
    End If
End Sub

Private Sub cmdOk_Click()
    SaveUpdateSetting chkEnabled.value
    SaveBetaUpdateSetting chkExperimental.value
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    CenterForm
End Sub

Private Sub UserForm_Initialize()
    lblUserAgent.Caption = GetUserAgent()
    chkEnabled.value = GetUpdateSetting()
    chkExperimental.value = GetBetaUpdateSetting()
    chkEnabled_Change
    
    AutoLayout
    CenterForm
End Sub

Private Sub AutoLayout()
    AutoFormat Me.Controls
    
    Me.Width = FormWidthUpdateSettings
    
    With lblDesc
        .Caption = "OpenSolver can automatically check for updates and let you know when a new version is available. " & _
                   "We only check for updates if OpenSolver is being actively used, and not more than once a day."
        .Left = FormMargin
        .Top = FormMargin
        AutoHeight lblDesc, Me.Width - 2 * FormMargin
    End With
    
    With chkEnabled
        .Width = lblDesc.Width
        .Left = lblDesc.Left
        .Top = Below(lblDesc)
    End With
    
    With chkExperimental
        .Width = lblDesc.Width
        .Left = lblDesc.Left
        .Top = Below(chkEnabled, False)
    End With
    
    With lblInfo
        .Caption = "Our update check sends anonymous version information which lets us collect statistics on the " & _
                   "Operating System, Excel, and OpenSolver versions being used. This helps ensure we are testing " & _
                   "OpenSolver on all popular platforms. If you enable update checks, the information sent for your installation would be:"
        .Left = lblDesc.Left
        .Top = Below(chkExperimental)
        AutoHeight lblInfo, lblDesc.Width
    End With
    
    With lblUserAgent
        .Left = lblDesc.Left
        .Top = Below(lblInfo)
        AutoHeight lblUserAgent, lblDesc.Width
        .TextAlign = fmTextAlignCenter
        .BackColor = FormBackColor
    End With
    
    With cmdCancel
        .Caption = "Cancel"
        .Width = FormButtonWidth
        .Left = LeftOfForm(Me.Width, .Width)
        .Top = Below(lblUserAgent)
        .Cancel = True
    End With
    
    With cmdOk
        .Caption = "OK"
        .Width = FormButtonWidth
        .Left = LeftOf(cmdCancel, .Width)
        .Top = cmdCancel.Top
    End With
        
    Me.Height = FormHeight(cmdOk)
    Me.Width = Me.Width + FormWindowMargin
    
    Me.BackColor = FormBackColor
    Me.Caption = "OpenSolver - Update Settings"
End Sub

Private Sub CenterForm()
    Me.Top = CenterFormTop(Me.Height)
    Me.Left = CenterFormLeft(Me.Width)
End Sub
