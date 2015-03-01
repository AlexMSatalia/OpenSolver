VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMessageBox 
   Caption         =   "OpenSolver - Message"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6120
   OleObjectBlob   =   "frmMessageBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If Mac Then
    Const FormWidthMessageBox = 400
#Else
    Const FormWidthMessageBox = 300
#End If

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub lblLink_Click()
    Call OpenURL(lblLink.ControlTipText)
End Sub

Private Sub txtMessage_Change()

End Sub

Private Sub UserForm_Activate()
    AutoLayout
End Sub

Private Sub AutoLayout()
    AutoFormat Me.Controls
    
    Me.width = FormWidthMessageBox
    
    With txtMessage
        .BackColor = FormBackColor
        .width = Me.width - 2 * FormMargin
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
        .width = Me.width - 2 * FormMargin
        .left = FormMargin
        .top = FormMargin
    End With
    
    With lblLink
        If .Caption = "" Then
            .Visible = False
        Else
            .width = Me.width
            .AutoSize = False
            .AutoSize = True
            .AutoSize = False
            .left = (Me.width - .width) / 2
        End If
        .top = txtMessage.top + txtMessage.height + FormSpacing
    End With
    
    With cmdOk
        .Caption = "OK"
        .top = IIf(lblLink.Visible, lblLink.height + FormSpacing, 0) + lblLink.top
        .width = FormButtonWidth
        .left = Me.width - FormMargin - .width
    End With
    
    Me.height = cmdOk.top + cmdOk.height + FormMargin + FormTitleHeight
    Me.width = Me.width + FormWindowMargin
    
    Me.BackColor = FormBackColor
End Sub
