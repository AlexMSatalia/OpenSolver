VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FSolveEngine 
   Caption         =   "Satalia Solve Engine running. . ."
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "FSolveEngine.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FSolveEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public shouldCancel As Boolean

#If Mac Then
    Const FormWidthSolveEngine = 365
#Else
    Const FormWidthSolveEngine = 220
#End If

Public Sub UpdateStatus(ByVal msg As String)
1         Me.lblStatusBar.Caption = msg
End Sub

Private Sub cmdCancel_Click()
1         Me.shouldCancel = True
2         Me.cmdCancel.Enabled = False
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

Public Sub UserForm_Activate()
1         Me.shouldCancel = False
          
          Dim errorString As String
2         SolverSolveEngine.SolveEngineFinalResponse = SolveOnSolveEngine( _
              SolverSolveEngine.SolveEngineLpModel, _
              SolverSolveEngine.SolveEngineLogPath, _
              errorString, _
              Me)
          
3         Me.Tag = errorString
          
4         Me.Hide
End Sub

Private Sub UserForm_Initialize()
1         AutoLayout
2         CenterForm
End Sub

Private Sub AutoLayout()
1         AutoFormat Me.Controls

2         Me.Width = FormWidthSolveEngine
          
3         With imgSatalia
4             .Left = (Me.Width - 2 * FormMargin - .Width) / 2 + FormMargin
5             .Top = FormMargin
              ' Mac doesn't show the image so hide it
6             If IsMac Then .Visible = False
7         End With
          
8         With lblStatusBar
9             .Caption = "OpenSolver is busy running your optimisation model..."
10            .Left = FormMargin
11            .Top = IIf(IsMac, FormMargin, Below(imgSatalia))
12            .Width = Me.Width - 2 * FormMargin
13            AutoHeight lblStatusBar, .Width, False
14            .Height = 3 * .Height
15        End With
          
16        With cmdCancel
17            .Caption = "Cancel"
18            .Width = FormButtonWidth
19            .Left = (lblStatusBar.Width - .Width) / 2 + lblStatusBar.Left
20            .Top = Below(lblStatusBar)
21            .Cancel = True
22        End With
          
23        Me.Height = FormHeight(cmdCancel)
24        Me.Width = Me.Width + FormWindowMargin
          
25        Me.BackColor = FormBackColor
26        Me.Caption = "OpenSolver - Optimisation Running"
End Sub

Private Sub CenterForm()
1         Me.Top = CenterFormTop(Me.Height)
2         Me.Left = CenterFormLeft(Me.Width)
End Sub
