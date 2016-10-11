VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FCallingNeos 
   Caption         =   "OpenSolver Optimisation Running"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   -945
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
1         Me.Hide
2         Me.Tag = "Cancelled"
End Sub

' Make the [x] hide the form rather than unload
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
          ' If CloseMode = vbFormControlMenu then we know the user
          ' clicked the [x] close button or Alt+F4 to close the form.
1         If CloseMode = vbFormControlMenu Then
2             cmdCancel_Click
3             Cancel = True
4         End If
End Sub

Private Sub UserForm_Activate()
1         CenterForm
          
          Dim message As String, errorString As String, result As String
2         message = SolverNeos.FinalMessage
3         errorString = vbNullString

4         result = SolveOnNeos(message, errorString, Me)

5         SolverNeos.NeosResult = result
6         Me.Tag = errorString
7         Me.Hide
End Sub

Private Sub UserForm_Initialize()
1        AutoLayout
2        CenterForm
End Sub

Private Sub AutoLayout()
1         AutoFormat Me.Controls
          
          ' Make the label wide enough so that the message is on one line, then use autosize to shrink the width.
2         With lblMessage
3             .Caption = "OpenSolver is busy running your optimisation model..."
4             .Left = FormMargin
5             .Top = FormMargin
6             AutoHeight lblMessage, FormWidthCallingNeos, True
7         End With
          
8         Me.Width = lblMessage.Width + 2 * FormMargin
          
9         With cmdCancel
10            .Caption = "Cancel"
11            .Width = FormButtonWidth
12            .Left = (lblMessage.Width - .Width) / 2 + lblMessage.Left
13            .Top = Below(lblMessage)
14            .Cancel = True
15        End With
          
16        Me.Height = FormHeight(cmdCancel)
17        Me.Width = Me.Width + FormWindowMargin
          
18        Me.BackColor = FormBackColor
19        Me.Caption = "OpenSolver - Optimisation Running"
End Sub

Private Sub CenterForm()
1         Me.Top = CenterFormTop(Me.Height)
2         Me.Left = CenterFormLeft(Me.Width)
End Sub
