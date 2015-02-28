VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNonlinear 
   Caption         =   "UserForm1"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11310
   OleObjectBlob   =   "frmNonlinear.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNonlinear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If Mac Then
    Const FormWidthNonLinear = 570
    Const MaxHeight = 350
#Else
    Const FormWidthNonLinear = 570
    Const MaxHeight = 250
#End If

Private Sub cmdContinue_Click()
3585      Me.Hide
End Sub

Public Sub SetLinearityResult(resultString As String, IsQuickCheck As Boolean)
    txtNonLinearInfo.Caption = resultString
    chkFullCheck.Visible = IsQuickCheck
    AutoLayout
End Sub

Private Sub AutoLayout()
    AutoFormat Me.Controls
    
    Me.width = FormWidthNonLinear
    
    With txtNonLinearInfo
        .left = FormMargin
        .top = FormMargin
        .width = Me.width - 2 * FormMargin
        .height = 20
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
        If .height > MaxHeight Then .height = MaxHeight
    End With
    
    With cmdContinue
        .left = txtNonLinearInfo.left
        .top = txtNonLinearInfo.height + txtNonLinearInfo.top + FormSpacing
        If chkFullCheck.Visible Then .top = .top + chkHighlight.height - .height / 2
        .width = FormButtonWidth
        .Caption = "Continue"
    End With
    
    With chkHighlight
        .left = cmdContinue.left + cmdContinue.width + FormSpacing
        .top = txtNonLinearInfo.height + txtNonLinearInfo.top + FormSpacing
        .width = Me.width - .left - FormMargin
        .Caption = "Highlight the nonlinearities"
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
    End With
    
    With chkFullCheck
        .left = chkHighlight.left
        .top = chkHighlight.height + chkHighlight.top
        .width = Me.width - .left - FormMargin
        .Caption = "Run a full linearity check. (This will destroy the current solution)"
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
    End With
    
    ' Adjust width to rightmost element
    Me.width = chkFullCheck.left + chkFullCheck.width
    If Me.width < txtNonLinearInfo.width + FormMargin Then Me.width = txtNonLinearInfo.width + FormMargin
    Me.width = Me.width + FormMargin + FormWindowMargin
    
    ' Adjust heights based on visible elements
    If chkFullCheck.Visible Then
        Me.height = chkFullCheck.top + chkFullCheck.height
    Else
        Me.height = cmdContinue.top + cmdContinue.height
    End If
    Me.height = Me.height + FormMargin + FormTitleHeight
    
    Me.BackColor = FormBackColor
    Me.Caption = "OpenSolver - Linearity Check"
End Sub
