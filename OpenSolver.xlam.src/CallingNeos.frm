VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CallingNeos 
   Caption         =   "OpenSolver Optimisation Running"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4560
   OleObjectBlob   =   "CallingNeos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CallingNeos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If Mac Then
    Const FormWidthCallingNeos = 350
#Else
    Const FormWidthCallingNeos = 240
#End If

Private Sub cmdCancel_Click()
    Me.Hide
    Me.Tag = "Cancelled"
End Sub

Private Sub UserForm_Initialize()
   AutoLayout
End Sub

Private Sub AutoLayout()
    AutoFormat Me.Controls
    
    ' Make the label wide enough so that the message is on one line, then use autosize to shrink the width.
    With lblMessage
        .Caption = "OpenSolver is busy running your optimisation model..."
        .left = FormMargin
        .top = FormMargin
        .width = FormWidthCallingNeos
        .AutoSize = False
        .AutoSize = True
        .AutoSize = False
    End With
    
    Me.width = lblMessage.width + 2 * FormMargin
    
    With cmdCancel
        .Caption = "Cancel"
        .width = FormButtonWidth
        .left = (lblMessage.width - .width) / 2 + lblMessage.left
        .top = lblMessage.top + lblMessage.height + FormSpacing * 2
    End With
    
    Me.height = cmdCancel.top + cmdCancel.height + FormSpacing + FormMargin + FormTitleHeight
    Me.width = Me.width + FormWindowMargin
    
    Me.BackColor = FormBackColor
    Me.Caption = "OpenSolver - Optimisation Running"
End Sub
