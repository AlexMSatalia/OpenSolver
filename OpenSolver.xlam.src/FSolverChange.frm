VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FSolverChange 
   Caption         =   "Choose Solver"
   ClientHeight    =   4648
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

Private Sub cboSolver_Change()
          Dim Solver As ISolver
4724      Set Solver = Solvers(cboSolver.ListIndex)
4725      lblDesc.Caption = Solver.Desc
          If TypeOf Solver Is ISolverNeos Then lblDesc.Caption = lblDesc.Caption & vbNewLine & vbNewLine & NeosAdditionalSolverText

          lblHyperlink.Caption = Solver.Link
                
          Dim errorString As String
4727      cmdOk.Enabled = SolverIsPresent(Solver, errorString:=errorString)
4732      lblError.Caption = errorString ' empty if no errors found

          AutoLayout
End Sub

Private Sub lblHyperlink_Click()
4733      OpenURL lblHyperlink.Caption
End Sub

Private Sub UserForm_Activate()
          CenterForm

4744      cboSolver.Clear
4745      cboSolver.MatchRequired = True
4746      cboSolver.Style = fmStyleDropDownList
          
          Dim ChosenSolver As String
          ChosenSolver = GetChosenSolver()
          
          Dim NumSolvers As Long
          NumSolvers = UBound(GetAvailableSolvers) - LBound(GetAvailableSolvers) + 1
          
          ReDim Solvers(0 To NumSolvers - 1)
          
          Dim Solver As Variant, SolverString As String, i As Long
          i = 0
4747      For Each Solver In GetAvailableSolvers()
              SolverString = CStr(Solver)
              Set Solvers(i) = CreateSolver(SolverString)
4748          cboSolver.AddItem Solvers(i).Title
              If Solvers(i).ShortName = ChosenSolver Then cboSolver.ListIndex = i
              i = i + 1
4749      Next Solver
End Sub

Private Sub cmdOk_Click()
         'Add the chosen solver as a hidden name in the workbook
4758      SetChosenSolver Solvers(cboSolver.ListIndex).ShortName
4763      Me.Hide
End Sub

Private Sub cmdCancel_Click()
4764      Me.Hide
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

Private Sub UserForm_Initialize()
    AutoLayout
    CenterForm
End Sub

Private Sub AutoLayout()
    AutoFormat Me.Controls
    
    Me.width = FormWidthSolverChange
    
    With lblChoose
        .left = FormMargin
        .top = FormMargin
        .width = Me.width - FormMargin * 2
        .Caption = "Choose a solver from the list below:"
    End With
    
    With cboSolver
        .left = lblChoose.left
        .top = Below(lblChoose, False)
        .width = lblChoose.width
    End With
    
    With lblDesc
        .left = lblChoose.left
        .top = Below(cboSolver)
        AutoHeight lblDesc, lblChoose.width
    End With
    
    With lblHyperlink
        .left = lblChoose.left
        .top = Below(lblDesc)
        AutoHeight lblHyperlink, lblChoose.width, True
    End With
    
    With lblError
        .Visible = Len(.Caption) <> 0
        .left = lblChoose.left
        .top = Below(lblHyperlink)
        AutoHeight lblError, lblChoose.width
    End With
    
    With cmdCancel
        .Caption = "Cancel"
        .width = FormButtonWidth
        .top = Below(IIf(lblError.Visible, lblError, lblHyperlink))
        .left = LeftOfForm(Me.width, .width)
    End With
    
    With cmdOk
        .Caption = "OK"
        .width = FormButtonWidth
        .top = cmdCancel.top
        .left = LeftOf(cmdCancel, .width)
    End With
    
    
    Me.height = FormHeight(cmdCancel)
    Me.width = Me.width + FormWindowMargin
    
    Me.BackColor = FormBackColor
    Me.Caption = "OpenSolver - Choose Solver"
End Sub

Private Sub CenterForm()
    Me.top = CenterFormTop(Me.height)
    Me.left = CenterFormLeft(Me.width)
End Sub
