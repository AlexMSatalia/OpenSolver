Attribute VB_Name = "OpenSolverQuickSolve"
Option Explicit

Public QuickSolver As COpenSolver

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
          If Not ReportError("OpenSolverQuickSolve", "SetQuickSolveParameterRange") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function
