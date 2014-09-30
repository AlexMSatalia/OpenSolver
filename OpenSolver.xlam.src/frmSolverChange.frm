VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSolverChange 
   Caption         =   "Choose Solver"
   ClientHeight    =   4485
   ClientLeft      =   42
   ClientTop       =   378
   ClientWidth     =   4816
   OleObjectBlob   =   "frmSolverChange.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSolverChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ChosenSolver As String
Dim Solvers As Collection

Private Sub cboSolver_Change()
4723      ChangeSolver Me
End Sub

Public Sub ChangeSolver(f As UserForm)
4724            ChosenSolver = ReverseSolverTitle(f.cboSolver.Text)
4725            f.lblSolver.Caption = SolverDesc(ChosenSolver)
4726            f.lblHyperLink = SolverLink(ChosenSolver)
                
                Dim errorString As String
4727            If SolverAvailable(ChosenSolver, errorString:=errorString) Then
4728                f.cmdOK.Enabled = True
4729            Else
4730                f.cmdOK.Enabled = False
4731            End If
4732            f.lblError.Caption = errorString ' empty if no errors found
End Sub

Private Sub lblHyperlink_Click()
4733      OpenURL lblHyperLink.Caption
End Sub

Private Sub UserForm_Activate()
4734      ActivateSolverChange Me
End Sub

Public Sub ActivateSolverChange(f As UserForm)
4735      Set Solvers = New Collection
4736      Solvers.Add "CBC"
4737      Solvers.Add "Gurobi"
          'Solvers.Add "Cplex"
4738      Solvers.Add "NeosCBC"
4739      Solvers.Add "Bonmin"
4740      Solvers.Add "Couenne"
4741      Solvers.Add "NOMAD"
4742      Solvers.Add "NeosBon"
4743      Solvers.Add "NeosCou"
          'Solvers.Add "PuLP"
          
4744      f.cboSolver.Clear
4745      f.cboSolver.MatchRequired = True
4746      f.cboSolver.Style = fmStyleDropDownList
          
          Dim Solver As Variant
4747      For Each Solver In Solvers
4748          f.cboSolver.AddItem SolverTitle(CStr(Solver))
4749      Next Solver

          Dim value As String
4750      If GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!OpenSolver_ChosenSolver", value) Then
4751          On Error GoTo setDefault
4752          f.cboSolver.Text = SolverTitle(value)
4753      Else
setDefault:
4754          f.cboSolver.Text = SolverTitle("CBC")
4755      End If
End Sub

Private Sub cmdOK_Click()
4756      SolverChangeConfirm Me
4757      Unload Me
End Sub

Public Sub SolverChangeConfirm(f As UserForm)
          'Add the chosen solver as a hidden name in the workbook
4758      Call SetNameOnSheet("OpenSolver_ChosenSolver", "=" & ChosenSolver)
#If Mac Then
4759      MacModel.lblSolver.Caption = "Current Solver Engine: " & UCase(left(ChosenSolver, 1)) & Mid(ChosenSolver, 2)
4760      frmModel.Disabler True, MacModel
#Else
4761      frmModel.lblSolver.Caption = "Current Solver Engine: " & UCase(left(ChosenSolver, 1)) & Mid(ChosenSolver, 2)
4762      frmModel.Disabler True, frmModel
#End If

4763      Unload Me
End Sub

Private Sub cmdCancel_Click()
4764      Unload Me
End Sub
