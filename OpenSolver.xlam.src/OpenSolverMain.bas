Attribute VB_Name = "OpenSolverMain"
Option Explicit

Public Const sOpenSolverVersion As String = "2.6.1"
Public Const sOpenSolverDate As String = "2015.02.15"

Dim OpenSolver As COpenSolver

Function RunOpenSolver(Optional SolveRelaxation As Boolean = False, Optional MinimiseUserInteraction As Boolean = False, Optional LinearityCheckOffset As Double = 0) As OpenSolverResult
          On Error GoTo ErrorHandler

2804      RunOpenSolver = OpenSolverResult.Unsolved
2805      Set OpenSolver = New COpenSolver

2806      OpenSolver.BuildModelFromSolverData LinearityCheckOffset, MinimiseUserInteraction, SolveRelaxation
          ' Only proceed with solve if nothing detected while building model
          If OpenSolver.SolveStatus = OpenSolverResult.Unsolved Then
              SolveModel OpenSolver, SolveRelaxation, MinimiseUserInteraction
          End If
          
          RunOpenSolver = OpenSolver.SolveStatus
          If Not MinimiseUserInteraction Then OpenSolver.ReportAnySolutionSubOptimality

ExitFunction:
          Set OpenSolver = Nothing    ' Free any OpenSolver memory used
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

Function SetQuickSolveParameterRange() As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

362       SetQuickSolveParameterRange = False
363       If Application.Workbooks.Count = 0 Then
364           Err.Raise OpenSolver_BuildError, Description:="No active workbook available"
366       End If
          
          ' Find the Parameter range
          Dim ParamRange As Range
375       Set ParamRange = GetQuickSolveParameters()
          
          ' Get a range from the user
          Dim NewRange As Range
377       On Error Resume Next
378       If ParamRange Is Nothing Then
379           Set NewRange = Application.InputBox(prompt:="Please select the 'parameter' cells that you will be changing between successsive solves of the model.", Type:=8, Title:="OpenSolver Quick Solve Parameters")
380       Else
381           Set NewRange = Application.InputBox(prompt:="Please select the 'parameter' cells that you will be changing between successsive solves of the model.", Type:=8, Default:=ParamRange.Address, Title:="OpenSolver Quick Solve Parameters")
382       End If
383       On Error GoTo ErrorHandler
          
384       If Not NewRange Is Nothing Then
385           If NewRange.Worksheet.Name <> ActiveSheet.Name Then
386               Err.Raise OpenSolver_BuildError, Description:="The parameter cells need to be on the current worksheet."
388           End If
389           SetQuickSolveParameters NewRange

              ' Return true as we have succeeded
393           SetQuickSolveParameterRange = True
394       End If

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverMain", "SetQuickSolveParameterRange") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function CheckModelHasParameterRange() As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

395       If Application.Workbooks.Count = 0 Then
396           Err.Raise OpenSolver_BuildError, Description:="No active workbook available"
398       End If
          
          ' Find the Parameter range
          Dim ParamRange As Range
408       Set ParamRange = GetQuickSolveParameters()
409       If ParamRange Is Nothing Then
411           Err.Raise OpenSolver_BuildError, Description:="No parameter range could be found on the worksheet. Please use the Initialize Quick Solve Parameters menu item to define the cells that you wish to change between successive OpenSolver solves. Note that changes to these cells must lead to changes in the underlying model's right hand side values for its constraints."
413       End If
406       CheckModelHasParameterRange = True

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverMain", "CheckModelHasParameterRange") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Sub InitializeQuickSolve()
          On Error GoTo ErrorHandler

          If Not CreateSolver(GetChosenSolver()).ModelType = Diff Then
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

Function RunQuickSolve(Optional MinimiseUserInteraction As Boolean = False) As OpenSolverResult
          On Error GoTo ErrorHandler

2840      If OpenSolver Is Nothing Then
2841          Err.Raise OpenSolver_SolveError, Description:="There is no model to solve, and so the quick solve cannot be completed. Please select the Initialize Quick Solve command."
2843      ElseIf OpenSolver.CanDoQuickSolveForActiveSheet Then    ' This will report any errors
2844          OpenSolver.DoQuickSolve MinimiseUserInteraction
              RunQuickSolve = OpenSolver.SolveStatus
2845      End If

          If Not MinimiseUserInteraction Then OpenSolver.ReportAnySolutionSubOptimality

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

