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
    Me.lblStatusBar.Caption = msg
End Sub

Private Sub cmdCancel_Click()
    Me.shouldCancel = True
    Me.cmdCancel.Enabled = False
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
    Me.shouldCancel = False
    
    Dim errorString As String
    SolverSolveEngine.SolveEngineFinalResponse = SolveOnSolveEngine( _
        SolverSolveEngine.SolveEngineLpModel, _
        SolverSolveEngine.SolveEngineLogPath, _
        errorString, _
        Me)
    
    Me.Tag = errorString
    
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    AutoLayout
    CenterForm
End Sub

Private Sub AutoLayout()
1         AutoFormat Me.Controls

          Me.Width = FormWidthSolveEngine
          
          With imgSatalia
              .Left = (Me.Width - 2 * FormMargin - .Width) / 2 + FormMargin
              .Top = FormMargin
              ' Mac doesn't show the image so hide it
              If IsMac Then .Visible = False
          End With
          
2         With lblStatusBar
3             .Caption = "OpenSolver is busy running your optimisation model..."
4             .Left = FormMargin
              .Top = IIf(IsMac, FormMargin, Below(imgSatalia))
6             .Width = Me.Width - 2 * FormMargin
              AutoHeight lblStatusBar, .Width, False
              .Height = 3 * .Height
7         End With
          
9         With cmdCancel
10            .Caption = "Cancel"
11            .Width = FormButtonWidth
12            .Left = (lblStatusBar.Width - .Width) / 2 + lblStatusBar.Left
13            .Top = Below(lblStatusBar)
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
