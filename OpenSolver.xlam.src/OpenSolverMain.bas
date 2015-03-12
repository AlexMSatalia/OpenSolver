Attribute VB_Name = "OpenSolverMain"
Option Explicit

Public Const sOpenSolverVersion As String = "2.6.1"
Public Const sOpenSolverDate As String = "2015.02.15"

Dim OpenSolver As COpenSolver

Function RunOpenSolver(Optional SolveRelaxation As Boolean = False, Optional MinimiseUserInteraction As Boolean = False, Optional LinearityCheckOffset As Double = 0) As OpenSolverResult
          On Error GoTo ErrorHandler

          'Save iterative calcalation state
          Dim oldIterationMode As Boolean
2803      oldIterationMode = Application.Iteration

2804      RunOpenSolver = OpenSolverResult.Unsolved
2805      Set OpenSolver = New COpenSolver
2806      OpenSolver.BuildModelFromSolverData LinearityCheckOffset, MinimiseUserInteraction, SolveRelaxation

          ' Run appropriate solve routine
          Dim OpenSolverParsed As COpenSolverParsed
2807      If UsesParsedModel(OpenSolver.Solver) Then
              Set OpenSolverParsed = New COpenSolverParsed
              OpenSolverParsed.SolveModel OpenSolver, SolveRelaxation, MinimiseUserInteraction
              RunOpenSolver = OpenSolver.SolveStatus
2809      Else
2810          RunOpenSolver = OpenSolver.SolveModel(SolveRelaxation, MinimiseUserInteraction)
2811      End If

          If Not MinimiseUserInteraction Then OpenSolver.ReportAnySolutionSubOptimality

ExitFunction:
          Set OpenSolver = Nothing    ' Free any OpenSolver memory used
          Set OpenSolverParsed = Nothing
          Application.Iteration = oldIterationMode
          Exit Function

ErrorHandler:
          ReportError "OpenSolverMain", "RunOpenSolver", True, MinimiseUserInteraction
          If OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError Then
              RunOpenSolver = AbortedThruUserAction
          Else
              RunOpenSolver = OpenSolverResult.ErrorOccurred
          End If
          GoTo ExitFunction
End Function

Sub InitializeQuickSolve()
          On Error GoTo ErrorHandler

          If UsesParsedModel(GetChosenSolver) Then
              Err.Raise OpenSolver_ModelError, Description:="The selected solver does not support QuickSolve"
          End If

2832      If CheckModelHasParameterRange Then
2833          Set OpenSolver = New COpenSolver
2834          OpenSolver.BuildModelFromSolverData
2835          OpenSolver.InitializeQuickSolve
2836      End If

ExitSub:
          Exit Sub

ErrorHandler:
          ReportError "OpenSolverMain", "InitializeQuickSolve", True
          GoTo ExitSub
End Sub

Function RunQuickSolve(Optional MinimiseUserInteraction As Boolean = False) As Long
          On Error GoTo ErrorHandler

2840      If OpenSolver Is Nothing Then
2841          MsgBox "Error: There is no model to solve, and so the quick solve cannot be completed. Please select the Initialize Quick Solve command.", , "OpenSolver" & sOpenSolverVersion & " Error"
              'MsgBoxEx "Help_Error: There is no model to solve, and so the quick solve cannot be completed. Please select the Initialize Quick Solve command."
2842          RunQuickSolve = OpenSolverResult.ErrorOccurred
2843      ElseIf OpenSolver.CanDoQuickSolveForActiveSheet Then    ' This will report any errors
2844          RunQuickSolve = OpenSolver.DoQuickSolve(MinimiseUserInteraction)
2845      End If

ExitFunction:
          Exit Function

ErrorHandler:
          ReportError "OpenSolverMain", "RunQuickSolve", True, MinimiseUserInteraction
          If OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError Then
              RunQuickSolve = AbortedThruUserAction
          Else
              RunQuickSolve = OpenSolverResult.ErrorOccurred
          End If
          GoTo ExitFunction
End Function

Sub ClearQuickSolve()
          Set OpenSolver = Nothing
End Sub

