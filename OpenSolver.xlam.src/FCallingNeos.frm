VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FCallingNeos 
   Caption         =   "OpenSolver Optimisation Running"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4560
   OleObjectBlob   =   "FCallingNeos.frx":0000
End
Attribute VB_Name = "FCallingNeos"
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

' Make the [x] hide the form rather than unload
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' If CloseMode = vbFormControlMenu then we know the user
    ' clicked the [x] close button or Alt+F4 to close the form.
    If CloseMode = vbFormControlMenu Then
        cmdCancel_Click
        Cancel = True
    End If
End Sub

Private Sub UserForm_Activate()
    CenterForm
    
    Dim message As String, errorString As String, result As String
    message = SolverNeos.FinalMessage
    errorString = ""

    result = SolveOnNeos(message, errorString, Me)

    SolverNeos.NeosResult = result
    Me.Tag = errorString
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
   AutoLayout
   CenterForm
End Sub

Private Sub AutoLayout()
    AutoFormat Me.Controls
    
    ' Make the label wide enough so that the message is on one line, then use autosize to shrink the width.
    With lblMessage
        .Caption = "OpenSolver is busy running your optimisation model..."
        .Left = FormMargin
        .Top = FormMargin
        AutoHeight lblMessage, FormWidthCallingNeos, True
    End With
    
    Me.Width = lblMessage.Width + 2 * FormMargin
    
    With cmdCancel
        .Caption = "Cancel"
        .Width = FormButtonWidth
        .Left = (lblMessage.Width - .Width) / 2 + lblMessage.Left
        .Top = Below(lblMessage)
    End With
    
    Me.Height = FormHeight(cmdCancel)
    Me.Width = Me.Width + FormWindowMargin
    
    Me.BackColor = FormBackColor
    Me.Caption = "OpenSolver - Optimisation Running"
End Sub

Private Sub CenterForm()
    Me.Top = CenterFormTop(Me.Height)
    Me.Left = CenterFormLeft(Me.Width)
End Sub
