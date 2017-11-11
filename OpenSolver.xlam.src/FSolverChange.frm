VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FSolverChange 
   Caption         =   "Choose Solver"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5010
   OleObjectBlob   =   "FSolverChange.frx":0000
End
Attribute VB_Name = "FSolverChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If Mac Then
    Const FormWidthSolverChange = 365
#Else
    Const FormWidthSolverChange = 255
#End If

Private Solvers() As ISolver
Private sheet As Worksheet

Private Sub cboSolver_Change()
          ' Make sure we don't get an error when esc is pressed
          ' The action we take is unimportant, we'll be exiting the form right after this
1         If cboSolver.ListIndex = -1 Then Exit Sub

          Dim Solver As ISolver
2         Set Solver = Solvers(cboSolver.ListIndex)
3         lblDesc.Caption = Solver.Desc
4         If TypeOf Solver Is ISolverNeos Then lblDesc.Caption = lblDesc.Caption & vbNewLine & vbNewLine & NeosAdditionalSolverText

5         lblHyperlink.Caption = Solver.Link
                
          Dim errorString As String
6         cmdOk.Enabled = SolverIsPresent(Solver, errorString:=errorString)
7         lblError.Caption = errorString ' empty if no errors found

8         AutoLayout
End Sub

Private Sub lblHyperlink_Click()
1         OpenURL lblHyperlink.Caption
End Sub

Private Sub UserForm_Activate()
1         CenterForm
          
2         GetActiveSheetIfMissing sheet

3         cboSolver.Clear
4         cboSolver.MatchRequired = True
5         cboSolver.Style = fmStyleDropDownList
          
          Dim ChosenSolver As String
6         ChosenSolver = GetChosenSolver(sheet)
          
          Dim NumSolvers As Long
7         NumSolvers = UBound(GetAvailableSolvers) - LBound(GetAvailableSolvers) + 1
          
8         ReDim Solvers(0 To NumSolvers - 1)
          
          Dim Solver As Variant, SolverString As String, i As Long
9         i = 0
10        For Each Solver In GetAvailableSolvers()
11            SolverString = CStr(Solver)
12            Set Solvers(i) = CreateSolver(SolverString)
13            cboSolver.AddItem Solvers(i).Title
14            If Solvers(i).ShortName = ChosenSolver Then cboSolver.ListIndex = i
15            i = i + 1
16        Next Solver
End Sub

Private Sub cmdOk_Click()
         'Add the chosen solver as a hidden name in the workbook
1         SetChosenSolver Solvers(cboSolver.ListIndex).ShortName, sheet
2         Me.Hide
End Sub

Private Sub cmdCancel_Click()
1         Me.Hide
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

Private Sub UserForm_Initialize()
1         AutoLayout
2         CenterForm
End Sub

Private Sub AutoLayout()
1         AutoFormat Me.Controls
          
2         Me.Width = FormWidthSolverChange
          
3         With lblChoose
4             .Left = FormMargin
5             .Top = FormMargin
6             .Width = Me.Width - FormMargin * 2
7             .Caption = "Choose a solver from the list below:"
8         End With
          
9         With cboSolver
10            .Left = lblChoose.Left
11            .Top = Below(lblChoose, False)
12            .Width = lblChoose.Width
              .ListRows = UBound(GetAvailableSolvers()) - LBound(GetAvailableSolvers()) + 1
13        End With
          
14        With lblDesc
15            .Left = lblChoose.Left
16            .Top = Below(cboSolver)
17            AutoHeight lblDesc, lblChoose.Width
18        End With
          
19        With lblHyperlink
20            .Left = lblChoose.Left
21            .Top = Below(lblDesc)
22            AutoHeight lblHyperlink, lblChoose.Width, True
23        End With
          
24        With lblError
25            .Visible = Len(.Caption) <> 0
26            .Left = lblChoose.Left
27            .Top = Below(lblHyperlink)
28            AutoHeight lblError, lblChoose.Width
29        End With
          
30        With cmdCancel
31            .Caption = "Cancel"
32            .Width = FormButtonWidth
33            .Top = Below(IIf(lblError.Visible, lblError, lblHyperlink))
34            .Left = LeftOfForm(Me.Width, .Width)
35            .Cancel = True
36        End With
          
37        With cmdOk
38            .Caption = "OK"
39            .Width = FormButtonWidth
40            .Top = cmdCancel.Top
41            .Left = LeftOf(cmdCancel, .Width)
42        End With
          
          
43        Me.Height = FormHeight(cmdCancel)
44        Me.Width = Me.Width + FormWindowMargin
          
45        Me.BackColor = FormBackColor
46        Me.Caption = "OpenSolver - Choose Solver"
End Sub

Private Sub CenterForm()
1         Me.Top = CenterFormTop(Me.Height)
2         Me.Left = CenterFormLeft(Me.Width)
End Sub
