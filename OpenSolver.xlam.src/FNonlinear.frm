VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FNonlinear 
   Caption         =   "UserForm1"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11310
   OleObjectBlob   =   "FNonlinear.frx":0000
End
Attribute VB_Name = "FNonlinear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If Mac Then
    Const FormWidthNonLinear = 570
    Const MinHeightNonLinear = 200
#Else
    Const FormWidthNonLinear = 570
    Const MinHeightNonLinear = 280
#End If

Private Sub cmdContinue_Click()
1         Me.Hide
End Sub

' Make the [x] hide the form rather than unload
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
          ' If CloseMode = vbFormControlMenu then we know the user
          ' clicked the [x] close button or Alt+F4 to close the form.
1         If CloseMode = vbFormControlMenu Then
2             cmdContinue_Click
3             Cancel = True
4         End If
End Sub

Public Sub SetLinearityResult(resultString As String, IsQuickCheck As Boolean)
1         With txtNonLinearInfo
2             .Locked = False
3             .Text = resultString
4             .Locked = True
5         End With
6         chkFullCheck.Visible = IsQuickCheck
7         AutoLayout
8         CenterForm
9         With txtNonLinearInfo
10            .SelStart = 0
11            .SetFocus
12        End With
End Sub

Private Sub UserForm_Activate()
1         CenterForm
End Sub

Private Sub AutoLayout()
1         AutoFormat Me.Controls
          
2         Me.Width = FormWidthNonLinear
          
3         With txtNonLinearInfo
4             .Left = FormMargin
5             .Top = FormMargin
6             AutoHeight txtNonLinearInfo, Me.Width - 2 * FormMargin, True
7             If .Height > MinHeightNonLinear Then
8                 .Height = MinHeightNonLinear
9                 .Width = .Width + 20  ' margin for scrollbar
10            End If
11            .BackColor = FormBackColor
12            .BackStyle = fmBackStyleTransparent
13        End With
          
14        With cmdContinue
15            .Left = txtNonLinearInfo.Left
16            .Top = Below(txtNonLinearInfo)
17            If chkFullCheck.Visible Then .Top = .Top + chkHighlight.Height - .Height / 2
18            .Width = FormButtonWidth
19            .Caption = "Continue"
20            .Cancel = True
21        End With
          
22        With chkHighlight
23            .Caption = "Highlight the nonlinearities"
24            .Left = RightOf(cmdContinue)
25            .Top = Below(txtNonLinearInfo)
26            AutoHeight chkHighlight, LeftOfForm(Me.Width, .Left), True
27        End With
          
28        With chkFullCheck
29            .Caption = "Run a full linearity check."
30            .Left = chkHighlight.Left
31            .Top = Below(chkHighlight, False)
32            AutoHeight chkFullCheck, LeftOfForm(Me.Width, .Left), True
33        End With
          
          ' Adjust width to rightmost element
34        Me.Width = RightOf(chkFullCheck, False)
35        If Me.Width < txtNonLinearInfo.Width + FormMargin Then Me.Width = txtNonLinearInfo.Width + FormMargin
36        Me.Width = Me.Width + FormMargin + FormWindowMargin
          
          ' Adjust heights based on visible elements
37        Me.Height = FormHeight(IIf(chkFullCheck.Visible, chkFullCheck, cmdContinue))
          
38        Me.BackColor = FormBackColor
39        Me.Caption = "OpenSolver - Linearity Check"
End Sub

Private Sub CenterForm()
1         Me.Top = CenterFormTop(Me.Height)
2         Me.Left = CenterFormLeft(Me.Width)
End Sub
