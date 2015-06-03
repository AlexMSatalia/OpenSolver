VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FCallingNeos 
   Caption         =   "OpenSolver Optimisation Running"
   ClientHeight    =   1834
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
        .left = FormMargin
        .top = FormMargin
        AutoHeight lblMessage, FormWidthCallingNeos, True
    End With
    
    Me.width = lblMessage.width + 2 * FormMargin
    
    With cmdCancel
        .Caption = "Cancel"
        .width = FormButtonWidth
        .left = (lblMessage.width - .width) / 2 + lblMessage.left
        .top = Below(lblMessage)
    End With
    
    Me.height = FormHeight(cmdCancel)
    Me.width = Me.width + FormWindowMargin
    
    Me.BackColor = FormBackColor
    Me.Caption = "OpenSolver - Optimisation Running"
End Sub

Private Sub CenterForm()
    Me.top = CenterFormTop(Me.height)
    Me.left = CenterFormLeft(Me.width)
End Sub
