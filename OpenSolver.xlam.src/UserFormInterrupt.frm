VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormInterrupt 
   Caption         =   "OpenSolver User Interrupt"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4710
   OleObjectBlob   =   "UserFormInterrupt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormInterrupt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#If Mac Then
    Const FormInterruptWidth = 312
#Else
    Const FormInterruptWidth = 240
#End If

Private Sub CommandButtonAbort_Click()
3540      Me.Tag = vbCancel
3541      Me.Hide
End Sub

Private Sub CommandButtonContinue_Click()
3542      Me.Tag = vbOK
3543      Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
3544      If CloseMode = vbFormControlMenu Then Cancel = True
End Sub

Private Sub UserForm_Initialize()
   Me.AutoLayout
End Sub

Sub AutoLayout()
    Me.width = FormInterruptWidth
    
    With TextBox1
        .Caption = "You have pressed the Escape key while the optimizer engine is running. Do you wish to stop solving this problem?"
        .Font.Name = FormFontName
        .Font.Size = FormFontSize
        .BackColor = FormBackColor
        .left = FormMargin
        .top = FormMargin
        .width = Me.width - 2 * FormMargin
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
    End With
    
    With CommandButtonContinue
        .Caption = "Continue"
        .Font.Name = FormFontName
        .Font.Size = FormFontSize
        .width = FormButtonWidth
        .height = FormButtonHeight
        .left = Me.width - FormMargin - .width
        .top = TextBox1.top + TextBox1.height + FormMargin
    End With
    
    With CommandButtonAbort
        .Caption = "Abort"
        .Font.Name = FormFontName
        .Font.Size = FormFontSize
        .width = FormButtonWidth
        .height = FormButtonHeight
        .left = CommandButtonContinue.left - FormMargin - .width
        .top = CommandButtonContinue.top
    End With
    
    Me.height = CommandButtonAbort.top + CommandButtonAbort.height + FormMargin + FormTitleHeight
    Me.width = Me.width + FormWindowMargin
    
    Me.BackColor = FormBackColor
    
    Me.Caption = "OpenSolver User Interrupt"
    
    
End Sub
