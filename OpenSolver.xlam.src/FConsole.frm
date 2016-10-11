VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FConsole 
   Caption         =   "OpenSolver - Solving Model"
   ClientHeight    =   5070
   ClientLeft      =   0
   ClientTop       =   -10125
   ClientWidth     =   7065
   OleObjectBlob   =   "FConsole.frx":0000
End
Attribute VB_Name = "FConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pCommand As String
Private pLogPath As String
Private pStartDir As String

Private pExitCode As Long
Private pConsoleOutput As String

#If Mac Then
    Const ConsoleWidth = 584
    Const ConsoleHeight = 500
#Else
    Const ConsoleWidth = 380
    Const ConsoleHeight = 400
#End If

Private Const MinWidth = ConsoleWidth
Private Const MinHeight = 200
Private ResizeStartX As Double
Private ResizeStartY As Double

Private Sub cmdCancel_Click()
1         ProcessAbortSignal
End Sub

Private Sub cmdOk_Click()
1         ProcessAbortSignal
End Sub

Private Sub txtConsole_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
          ' Override any escape keypress for the textbox so it doesn't clear the text
1         If KeyCode = 27 Then
2             KeyCode = 0
3             ProcessAbortSignal
4         End If
End Sub

' Make the [x] hide the form rather than unload
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
          ' If CloseMode = vbFormControlMenu then we know the user
          ' clicked the [x] close button or Alt+F4 to close the form.
1         If CloseMode = vbFormControlMenu Then
2             ProcessAbortSignal
3             Cancel = True
4         End If
End Sub

Public Sub SetInput(Command As String, LogPath As String, StartDir As String)
1         pCommand = Command
2         pLogPath = LogPath
3         pStartDir = StartDir
End Sub

Public Sub GetOutput(ByRef ExitCode As Long, ByRef ConsoleOutput As String)
1         ExitCode = pExitCode
2         ConsoleOutput = pConsoleOutput
End Sub

Public Sub AppendText(NewText As String)
1         If Len(NewText) > 0 Then
2             With Me.txtConsole
3                 .Locked = False
4                 .Text = .Text & NewText
5                 .Locked = True
6             End With
7         End If
8         UpdateElapsedTime
End Sub

Private Sub UpdateElapsedTime()
1         Me.lblElapsed.Caption = "Elapsed Time: " & Int(Timer() - OpenSolverExternalCommand.StartTime) & "s"
End Sub

Public Sub MarkCompleted()
          Dim message As String
1         If Me.Tag = "Cancelled" Then
2             message = "Process cancelled."
3         ElseIf pExitCode <> 0 Then
4             message = "Process exited abnormally with exit code " & pExitCode & "."
5         Else
6             message = "Process completed successfully."
7         End If
8         Me.AppendText vbNewLine & vbNewLine & message
          
          ' Scroll to bottom
9         Me.txtConsole.SetFocus
          
10        cmdCancel.Enabled = False
11        cmdOk.Enabled = True
12        cmdOk.SetFocus
End Sub

Private Sub ProcessAbortSignal()
1         If cmdCancel.Enabled Then
2             Me.Tag = "Cancelled"
3         Else
4             Me.Hide
5         End If
End Sub

Private Sub UserForm_Activate()
1         On Error GoTo ErrorHandler  ' Don't let an error propogate out of the execution
2         pConsoleOutput = ExecConsole(Me, pCommand, pLogPath, pStartDir, pExitCode)
3         Exit Sub
          
ErrorHandler:
4         If OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError Then
5             Me.Tag = "Aborted"
6         Else
7             Me.Tag = OpenSolverErrorHandler.ErrMsg
8         End If
End Sub

Private Sub UserForm_Initialize()
1        AutoLayout
2        CenterForm
End Sub

Private Sub AutoLayout()
1         AutoFormat Me.Controls
          
2         With Me.txtConsole
        #If Mac Then
3                 .Font.Name = "Menlo Regular"
        #Else
4                 .Font.Name = "Consolas"
        #End If
5             .ForeColor = &HFFFFFF
6             .BackColor = &H0
7             .MultiLine = True
8             .ScrollBars = fmScrollBarsVertical
9             .SpecialEffect = fmSpecialEffectEtched
10            .Width = ConsoleWidth
11            .Height = ConsoleHeight
12            .Top = FormMargin
13            .Left = FormMargin
14        End With
          
15        With Me.cmdCancel
16            .Caption = "Cancel"
17            .Width = FormButtonWidth
18            .Cancel = True
19            .Enabled = True
20        End With
          
21        With Me.cmdOk
22            .Caption = "OK"
23            .Width = FormButtonWidth
24            .Cancel = True
25            .Enabled = False
26        End With
          
          ' Make the label wide enough so that the message is on one line, then use autosize to shrink the width.
27        With Me.lblElapsed
28            .Caption = "OpenSolver is busy running your optimisation model..."
29            .Left = FormMargin
30        End With
          
          ' Add resizer
31        With lblResizer
        #If Mac Then
                  ' Mac labels don't fire MouseMove events correctly
32                .Visible = False
        #End If
33            .Caption = "o"
34            With .Font
35                .Name = "Marlett"
36                .Charset = 2
37                .Size = 10
38            End With
39            .AutoSize = True
40            .MousePointer = fmMousePointerSizeNWSE
41            .BackStyle = fmBackStyleTransparent
42        End With
          
          ' Set the positions of the form
43        UpdateLayout
          
44        Me.BackColor = FormBackColor
45        Me.Caption = "OpenSolver - Optimisation Running"
End Sub

Private Sub CenterForm()
1         Me.Top = CenterFormTop(Me.Height)
2         Me.Left = CenterFormLeft(Me.Width)
End Sub

Private Sub UpdateLayout(Optional ChangeX As Single = 0, Optional ChangeY As Single = 0)
          Dim NewWidth As Double, NewHeight As Double
1         NewWidth = Max(txtConsole.Width + ChangeX, MinWidth)
2         NewHeight = Max(txtConsole.Height + ChangeY, MinHeight)
          
          ' Update based on new width
3         txtConsole.Width = NewWidth
4         Me.Width = Me.txtConsole.Width + 2 * FormMargin
5         Me.cmdCancel.Left = LeftOfForm(Me.Width, Me.cmdCancel.Width) - 1  ' To account for etched effect on textbox
6         Me.cmdOk.Left = LeftOf(cmdCancel, Me.cmdOk.Width)
7         Me.lblElapsed.Width = Me.cmdOk.Left - Me.txtConsole.Left - FormMargin
8         AutoHeight Me.lblElapsed, Me.lblElapsed.Width
9         Me.lblResizer.Left = Me.Width - Me.lblResizer.Width
10        Me.Width = Me.Width + FormWindowMargin
          
          ' Update based on new height
11        txtConsole.Height = NewHeight
12        Me.cmdCancel.Top = Below(Me.txtConsole)
13        Me.cmdOk.Top = Me.cmdCancel.Top
14        Me.lblElapsed.Top = Me.cmdCancel.Top + (Me.cmdCancel.Height - Me.lblElapsed.Height) / 2
15        Me.Height = FormHeight(Me.cmdCancel)
16        Me.lblResizer.Top = Me.InsideHeight - Me.lblResizer.Height
End Sub

Private Sub lblResizer_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
1         If Button = 1 Then
2             ResizeStartX = X
3             ResizeStartY = Y
4         End If
End Sub

Private Sub lblResizer_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
1         If Button = 1 Then
        #If Mac Then
                  ' Mac reports delta already
2                 UpdateLayout X, Y
        #Else
3                 UpdateLayout X - ResizeStartX, Y - ResizeStartY
        #End If
4         End If
End Sub
