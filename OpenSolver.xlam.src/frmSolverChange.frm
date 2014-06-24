VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSolverChange 
   Caption         =   "Choose Solver"
   ClientHeight    =   4050
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

Private Sub cboSolver_Change()
          If cboSolver.Text Like "NEOS using COIN*" Then
              lblSolver.Caption = "The NEOS (Network-Enabled Optimization System) Server is a free Internet-based service for solving optimization problems. The COIN Branch and Cut solver (CBC) is the default solver for OpenSolver and is an open-source mixed-integer program (MIP) solver written in C++. CBC is an active open-source project led by John Forrest at www.coin-or.org."
              lblHyperlink = "http://www.neos-server.org/neos/solvers/milp:Cbc/AMPL.html"
              ChosenSolver = "NeosCBC"
          ElseIf cboSolver.Text Like "NEOS using Couenne*" Then
              lblSolver.Caption = "The NEOS (Network-Enabled Optimization System) Server is a free Internet-based service for solving optimization problems. Couenne (Convex Over and Under ENvelopes for Nonlinear Estimation) is a branch & bound algorithm to solve Mixed-Integer Nonlinear Programming (MINLP) problems of sepcific forms."
              lblHyperlink = "https://projects.coin-or.org/Couenne"
              ChosenSolver = "NeosCou"
          ElseIf cboSolver.Text Like "NEOS using Bonmin*" Then
              lblSolver.Caption = "The NEOS (Network-Enabled Optimization System) Server is a free Internet-based service for solving optimization problems. Bonmin (Basic Open-source Nonlinear Mixed INteger programming) is an experimental open-source C++ code for solving general MINLP (Mixed Integer NonLinear Programming) problems."
              lblHyperlink = "https://projects.coin-or.org/Bonmin"
              ChosenSolver = "NeosBon"
          'ElseIf cboSolver.Text Like "PuLP*" Then
          '    lblSolver.Caption = "PuLP is an LP modeler written in python. PuLP can generate MPS or LP filesand call GLPK, COIN CLP/CBC, CPLEX, and GUROBI to solve linearproblems."
          '    lblHyperlink = "https://pypi.python.org/pypi/PuLP"
          '    ChosenSolver = "PuLP"
47710     ElseIf cboSolver.Text Like "*CBC*" Then
47720         lblSolver.Caption = "The COIN Branch and Cut solver (CBC) is the default solver for OpenSolver and is an open-source mixed-integer program (MIP) solver written in C++. CBC is an active open-source project led by John Forrest at www.coin-or.org."
47730         lblHyperlink = "http://www.coin-or.org/Cbc/cbcuserguide.html"
47740         ChosenSolver = "CBC"
47750     ElseIf cboSolver.Text Like "Gurobi*" Then
47760         lblSolver.Caption = "Gurobi is a solver for linear programming (LP), quadratic and quadratically constrained programming (QP and QCP), and mixed-integer programming (MILP, MIQP, and MIQCP). It requires the user to download and install a version of the Gurobi and to have GurobiOSRun.py in the OpenSolver directory."
47770         lblHyperlink = "http://www.gurobi.com/resources/documentation"
47780         ChosenSolver = "Gurobi"
47790     ElseIf cboSolver.Text Like "NOMAD*" Then
47800         lblSolver.Caption = "Nomad (Nonsmooth Optimization by Mesh Adaptive Direct search) is a C++ implementation of the Mesh Adaptive Direct Search (Mads) algorithm that solves non-linear problems. It works by updating the values on the sheet and passing them to the C++ solver. Like many non-linear solvers NOMAD cannot guarantee optimality of its solutions."
47810         lblHyperlink = "http://www.gerad.ca/nomad/Project/Home.html"
47820         ChosenSolver = "NOMAD"
47830     End If
End Sub

Private Sub lblHyperlink_Click()
          Dim link As String
47840     link = lblHyperlink.Caption
47850     On Error GoTo NoCanDo
47860     ActiveWorkbook.FollowHyperlink Address:=link, NewWindow:=True
47870     Exit Sub
NoCanDo:
47880     MsgBox "Cannot open " & link
End Sub

Private Sub UserForm_Activate()
47890     cboSolver.Clear
47900     cboSolver.AddItem "COIN-OR CBC (Linear Solver)"
          cboSolver.AddItem "NEOS using COIN-OR CBC (Linear Solver)"
          'cboSolver.AddItem "Cplex"
47910     cboSolver.AddItem "Gurobi (Linear Solver)"
47920     cboSolver.AddItem "NOMAD (Non-linear Solver)"
          cboSolver.AddItem "NEOS using Couenne (Non-linear Solver)"
          cboSolver.AddItem "NEOS using Bonmin (Non-linear Solver)"
          'cboSolver.AddItem "PuLP (Non-linear Solver)"
47930     cboSolver.MatchRequired = True
47940     cboSolver.Style = fmStyleDropDownList
          Dim value As String
47950     If GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!OpenSolver_ChosenSolver", value) Then
              On Error GoTo setDefault
              If value Like "NeosCBC" Then
                  cboSolver.Text = "NEOS using COIN-OR CBC (Linear Solver)"
              ElseIf value Like "NeosCou" Then
                  cboSolver.Text = "NEOS using Couenne (Non-linear Solver)"
              ElseIf value Like "NeosBon" Then
                  cboSolver.Text = "NEOS using Bonmin (Non-linear Solver)"
              'ElseIf value Like "PuLP" Then
              '   cboSolver.Text = "PuLP (Non-linear Solver)"
47970         ElseIf value Like "CBC" Or value Like "cbc" Then
47980             cboSolver.Text = "COIN-OR CBC (Linear Solver)"
47990         ElseIf value Like "*urobi" Then
48000             cboSolver.Text = "Gurobi (Linear Solver)"
48010         ElseIf value = "NOMAD" Then
48020             cboSolver.Text = "NOMAD (Non-linear Solver)"
48030         Else: GoTo setDefault
48040         End If
48050     Else
setDefault:
48060         cboSolver.Text = "COIN-OR CBC (Linear Solver)"
48070     End If
48080     Exit Sub
End Sub

Private Sub cmdOK_Click()
          'Add the chosen solver as a hidden name in the workbook
48090     Call SetNameOnSheet("OpenSolver_ChosenSolver", "=" & ChosenSolver)
48100     frmModel.lblSolver.Caption = "Current Solver Engine: " & UCase(left(ChosenSolver, 1)) & Mid(ChosenSolver, 2)
48110     frmModel.Disabler (True)
          'Me.Hide
48120     Unload Me
End Sub

Private Sub cmdCancel_Click()
48130     Unload Me
End Sub
