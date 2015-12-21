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
3585      Me.Hide
End Sub

' Make the [x] hide the form rather than unload
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' If CloseMode = vbFormControlMenu then we know the user
    ' clicked the [x] close button or Alt+F4 to close the form.
    If CloseMode = vbFormControlMenu Then
        cmdContinue_Click
        Cancel = True
    End If
End Sub

Public Sub SetLinearityResult(resultString As String, IsQuickCheck As Boolean)
    With txtNonLinearInfo
        .Locked = False
        .Text = resultString
        .Locked = True
    End With
    chkFullCheck.Visible = IsQuickCheck
    AutoLayout
    CenterForm
    With txtNonLinearInfo
        .SelStart = 0
        .SetFocus
    End With
End Sub

Private Sub UserForm_Activate()
    CenterForm
End Sub

Private Sub AutoLayout()
    AutoFormat Me.Controls
    
    Me.Width = FormWidthNonLinear
    
    With txtNonLinearInfo
        .Left = FormMargin
        .Top = FormMargin
        AutoHeight txtNonLinearInfo, Me.Width - 2 * FormMargin, True
        If .Height > MinHeightNonLinear Then
            .Height = MinHeightNonLinear
            .Width = .Width + 20  ' margin for scrollbar
        End If
    End With
    
    With cmdContinue
        .Left = txtNonLinearInfo.Left
        .Top = Below(txtNonLinearInfo)
        If chkFullCheck.Visible Then .Top = .Top + chkHighlight.Height - .Height / 2
        .Width = FormButtonWidth
        .Caption = "Continue"
        .Cancel = True
    End With
    
    With chkHighlight
        .Caption = "Highlight the nonlinearities"
        .Left = RightOf(cmdContinue)
        .Top = Below(txtNonLinearInfo)
        AutoHeight chkHighlight, LeftOfForm(Me.Width, .Left), True
    End With
    
    With chkFullCheck
        .Caption = "Run a full linearity check."
        .Left = chkHighlight.Left
        .Top = Below(chkHighlight, False)
        AutoHeight chkFullCheck, LeftOfForm(Me.Width, .Left), True
    End With
    
    ' Adjust width to rightmost element
    Me.Width = RightOf(chkFullCheck, False)
    If Me.Width < txtNonLinearInfo.Width + FormMargin Then Me.Width = txtNonLinearInfo.Width + FormMargin
    Me.Width = Me.Width + FormMargin + FormWindowMargin
    
    ' Adjust heights based on visible elements
    Me.Height = FormHeight(IIf(chkFullCheck.Visible, chkFullCheck, cmdContinue))
    
    Me.BackColor = FormBackColor
    Me.Caption = "OpenSolver - Linearity Check"
End Sub

Private Sub CenterForm()
    Me.Top = CenterFormTop(Me.Height)
    Me.Left = CenterFormLeft(Me.Width)
End Sub
