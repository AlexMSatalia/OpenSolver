VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSolverChange 
   Caption         =   "Choose Solver"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3990
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
          ChosenSolver = ReverseSolverTitle(cboSolver.Text)
          lblSolver.Caption = SolverDesc(ChosenSolver)
          lblHyperlink = SolverLink(ChosenSolver)
          
          Dim errorString As String
          If SolverAvailable(ChosenSolver, errorString:=errorString) Then
              cmdOK.Enabled = True
          Else
              cmdOK.Enabled = False
          End If
          lblError.Caption = errorString ' empty if no errors found
End Sub

Private Sub lblHyperlink_Click()
          Dim link As String
47840     link = lblHyperlink.Caption
47850     OpenURL link
End Sub

Private Sub UserForm_Activate()
          Set Solvers = New Collection
          Solvers.Add "CBC"
          Solvers.Add "Gurobi"
          'Solvers.Add "Cplex"
          Solvers.Add "NeosCBC"
          Solvers.Add "NOMAD"
          Solvers.Add "NeosBon"
          Solvers.Add "NeosCou"
          'Solvers.Add "PuLP"
          
47890     cboSolver.Clear
47930     cboSolver.MatchRequired = True
47940     cboSolver.Style = fmStyleDropDownList
          
          Dim Solver As Variant
          For Each Solver In Solvers
              cboSolver.AddItem SolverTitle(CStr(Solver))
          Next Solver

          Dim value As String
47950     If GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!OpenSolver_ChosenSolver", value) Then
              On Error GoTo setDefault
              cboSolver.Text = SolverTitle(value)
48050     Else
setDefault:
48060         cboSolver.Text = SolverTitle("CBC")
48070     End If
End Sub

Private Sub cmdOK_Click()
          'Add the chosen solver as a hidden name in the workbook
48090     Call SetNameOnSheet("OpenSolver_ChosenSolver", "=" & ChosenSolver)
48100     frmModel.lblSolver.Caption = "Current Solver Engine: " & UCase(left(ChosenSolver, 1)) & Mid(ChosenSolver, 2)
48110     frmModel.Disabler (True)
48120     Unload Me
End Sub

Private Sub cmdCancel_Click()
48130     Unload Me
End Sub
