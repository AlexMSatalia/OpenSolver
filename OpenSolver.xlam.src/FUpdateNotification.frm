VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FUpdateNotification 
   Caption         =   "OpenSolver - Update Available"
   ClientHeight    =   3165
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
1         Me.Hide
End Sub

' Make the [x] hide the form rather than unload
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
          ' If CloseMode = vbFormControlMenu then we know the user
          ' clicked the [x] close button or Alt+F4 to close the form.
1         If CloseMode = vbFormControlMenu Then
2             cmdOk_Click
3             Cancel = True
4         End If
End Sub

Private Sub cmdSettings_Click()
          Dim frmUpdateSettings As FUpdateSettings
1         Set frmUpdateSettings = New FUpdateSettings
2         frmUpdateSettings.Show
3         Unload frmUpdateSettings
End Sub

Private Sub lblLink_Click()
1         OpenURL lblLink.Caption
End Sub

Private Sub UserForm_Activate()
1         CenterForm
End Sub

Private Sub UserForm_Initialize()
1         AutoLayout
2         CenterForm
End Sub

Private Sub AutoLayout()
1         AutoFormat Me.Controls
          
2         Me.Width = FormWidthUpdateNotification
          
3         With lblDesc
4             .Caption = "A newer version of OpenSolver is available. Please follow the link below for more information and to download the update:"
5             .Left = FormMargin
6             .Top = FormMargin
7             AutoHeight lblDesc, Me.Width - 2 * FormMargin
8         End With
          
9         With lblLatestVersion
10            .Width = lblDesc.Width / 2
11            .Left = lblDesc.Left
12            .Top = Below(lblDesc)
13        End With
          
14        With lblCurrentVersion
15            .Width = lblLatestVersion.Width
16            .Left = RightOf(lblLatestVersion, False)
17            .Top = lblLatestVersion.Top
18        End With
          
19        With lblLink
20            .Caption = "http://OpenSolver.org"
21            .ForeColor = FormLinkColor
22            .Font.Underline = True
23            .Left = lblDesc.Left
24            .Top = Below(lblLatestVersion, False)
25            .Width = lblDesc.Width
26            .TextAlign = fmTextAlignCenter
27        End With
          
28        With cmdSettings
29            .Caption = "Update Settings..."
30            .Width = (lblDesc.Width - FormSpacing) / 2
31            .Left = lblDesc.Left
32            .Top = Below(lblLink)
33        End With
          
34        With cmdOk
35            .Caption = "OK"
36            .Left = RightOf(cmdSettings)
37            .Width = cmdSettings.Width
38            .Top = cmdSettings.Top
39            .Cancel = True
40        End With
              
41        Me.Height = FormHeight(cmdOk)
42        Me.Width = Me.Width + FormWindowMargin
          
43        Me.BackColor = FormBackColor
44        Me.Caption = "OpenSolver - Update Available"
End Sub

Private Sub CenterForm()
1         Me.Top = CenterFormTop(Me.Height)
2         Me.Left = CenterFormLeft(Me.Width)
End Sub

Sub ShowUpdate(LatestVersion As String)
1         lblLatestVersion.Caption = "Latest version: " & LatestVersion
2         lblCurrentVersion.Caption = "Current version: " & sOpenSolverVersion
3         Me.Show
End Sub
