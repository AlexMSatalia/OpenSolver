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
#Else
    Const FormWidthNonLinear = 570
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
    txtNonLinearInfo.Caption = resultString
    chkFullCheck.Visible = IsQuickCheck
    AutoLayout
    CenterForm
End Sub

Private Sub UserForm_Activate()
    CenterForm
End Sub

Private Sub AutoLayout()
    AutoFormat Me.Controls
    
    Me.width = FormWidthNonLinear
    
    With txtNonLinearInfo
        .left = FormMargin
        .top = FormMargin
        AutoHeight txtNonLinearInfo, Me.width - 2 * FormMargin, True
    End With
    
    With cmdContinue
        .left = txtNonLinearInfo.left
        .top = Below(txtNonLinearInfo)
        If chkFullCheck.Visible Then .top = .top + chkHighlight.height - .height / 2
        .width = FormButtonWidth
        .Caption = "Continue"
    End With
    
    With chkHighlight
        .Caption = "Highlight the nonlinearities"
        .left = RightOf(cmdContinue)
        .top = Below(txtNonLinearInfo)
        AutoHeight chkHighlight, LeftOfForm(Me.width, .left), True
    End With
    
    With chkFullCheck
        .Caption = "Run a full linearity check. (This will destroy the current solution)"
        .left = chkHighlight.left
        .top = Below(chkHighlight, False)
        AutoHeight chkFullCheck, LeftOfForm(Me.width, .left), True
    End With
    
    ' Adjust width to rightmost element
    Me.width = RightOf(chkFullCheck, False)
    If Me.width < txtNonLinearInfo.width + FormMargin Then Me.width = txtNonLinearInfo.width + FormMargin
    Me.width = Me.width + FormMargin + FormWindowMargin
    
    ' Adjust heights based on visible elements
    Me.height = FormHeight(IIf(chkFullCheck.Visible, chkFullCheck, cmdContinue))
    
    Me.BackColor = FormBackColor
    Me.Caption = "OpenSolver - Linearity Check"
End Sub

Private Sub CenterForm()
    Me.top = CenterFormTop(Me.height)
    Me.left = CenterFormLeft(Me.width)
End Sub
