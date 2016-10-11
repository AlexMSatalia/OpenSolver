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
1         chkExperimental.Enabled = chkEnabled.value
End Sub

Private Sub cmdCancel_Click()
1         Me.Hide
End Sub

' Make the [x] hide the form rather than unload
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
          ' If CloseMode = vbFormControlMenu then we know the user
          ' clicked the [x] close button or Alt+F4 to close the form.
1         If CloseMode = vbFormControlMenu Then
2             cmdCancel_Click
3             Cancel = True
4         End If
End Sub

Private Sub cmdOk_Click()
1         SaveUpdateSetting chkEnabled.value
2         SaveBetaUpdateSetting chkExperimental.value
3         Me.Hide
End Sub

Private Sub UserForm_Activate()
1         CenterForm
End Sub

Private Sub UserForm_Initialize()
1         lblUserAgent.Caption = GetUserAgent()
2         chkEnabled.value = GetUpdateSetting()
3         chkExperimental.value = GetBetaUpdateSetting()
4         chkEnabled_Change
          
5         AutoLayout
6         CenterForm
End Sub

Private Sub AutoLayout()
1         AutoFormat Me.Controls
          
2         Me.Width = FormWidthUpdateSettings
          
3         With lblDesc
4             .Caption = "OpenSolver can automatically check for updates and let you know when a new version is available. " & _
                         "We only check for updates if OpenSolver is being actively used, and not more than once a day."
5             .Left = FormMargin
6             .Top = FormMargin
7             AutoHeight lblDesc, Me.Width - 2 * FormMargin
8         End With
          
9         With chkEnabled
10            .Width = lblDesc.Width
11            .Left = lblDesc.Left
12            .Top = Below(lblDesc)
13        End With
          
14        With chkExperimental
15            .Width = lblDesc.Width
16            .Left = lblDesc.Left
17            .Top = Below(chkEnabled, False)
18        End With
          
19        With lblInfo
20            .Caption = "Our update check sends anonymous version information which lets us collect statistics on the " & _
                         "Operating System, Excel, and OpenSolver versions being used. This helps ensure we are testing " & _
                         "OpenSolver on all popular platforms. If you enable update checks, the information sent for your installation would be:"
21            .Left = lblDesc.Left
22            .Top = Below(chkExperimental)
23            AutoHeight lblInfo, lblDesc.Width
24        End With
          
25        With lblUserAgent
26            .Left = lblDesc.Left
27            .Top = Below(lblInfo)
28            AutoHeight lblUserAgent, lblDesc.Width
29            .TextAlign = fmTextAlignCenter
30            .BackColor = FormBackColor
31        End With
          
32        With cmdCancel
33            .Caption = "Cancel"
34            .Width = FormButtonWidth
35            .Left = LeftOfForm(Me.Width, .Width)
36            .Top = Below(lblUserAgent)
37            .Cancel = True
38        End With
          
39        With cmdOk
40            .Caption = "OK"
41            .Width = FormButtonWidth
42            .Left = LeftOf(cmdCancel, .Width)
43            .Top = cmdCancel.Top
44        End With
              
45        Me.Height = FormHeight(cmdOk)
46        Me.Width = Me.Width + FormWindowMargin
          
47        Me.BackColor = FormBackColor
48        Me.Caption = "OpenSolver - Update Settings"
End Sub

Private Sub CenterForm()
1         Me.Top = CenterFormTop(Me.Height)
2         Me.Left = CenterFormLeft(Me.Width)
End Sub
