VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInterrupt 
   Caption         =   "OpenSolver User Interrupt"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4710
   OleObjectBlob   =   "frmInterrupt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInterrupt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If Mac Then
    Const FormWidthInterrupt = 312
#Else
    Const FormWidthInterrupt = 240
#End If

Private Sub cmdAbort_Click()
3540      Me.Tag = vbCancel
3541      Me.Hide
End Sub

Private Sub cmdContinue_Click()
3542      Me.Tag = vbOK
3543      Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
3544      If CloseMode = vbFormControlMenu Then Cancel = True
End Sub

Private Sub UserForm_Initialize()
   AutoLayout
End Sub

Private Sub AutoLayout()
    AutoFormat Me.Controls

    Me.width = FormWidthInterrupt
    
    With lblMessage
        .Caption = "You have pressed the Escape key while the optimizer engine is running. Do you wish to stop solving this problem?"
        .left = FormMargin
        .top = FormMargin
        .width = Me.width - 2 * FormMargin
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
        .width = Me.width - 2 * FormMargin
    End With
    
    With cmdContinue
        .Caption = "Continue"
        .width = FormButtonWidth
        .left = Me.width - FormMargin - .width
        .top = lblMessage.top + lblMessage.height + FormSpacing
    End With
    
    With cmdAbort
        .Caption = "Abort"
        .width = cmdContinue.width
        .left = cmdContinue.left - FormSpacing - .width
        .top = cmdContinue.top
    End With
    
    Me.height = cmdAbort.top + cmdAbort.height + FormMargin + FormTitleHeight
    Me.width = Me.width + FormWindowMargin
    
    Me.BackColor = FormBackColor
    Me.Caption = "OpenSolver - User Interrupt"
End Sub
