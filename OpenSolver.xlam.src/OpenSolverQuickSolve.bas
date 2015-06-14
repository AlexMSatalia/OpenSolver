Attribute VB_Name = "OpenSolverQuickSolve"
Option Explicit

Public QuickSolver As COpenSolver

Function SetQuickSolveParameterRange() As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

362       SetQuickSolveParameterRange = False
          
          Dim sheet As Worksheet
          GetActiveSheetIfMissing sheet
          
          ' Find the Parameter range
          Dim ParamRange As Range
375       Set ParamRange = GetQuickSolveParameters(sheet)
          
          ' Get a range from the user
          Dim DefaultValue As String
          Dim NewRange As Range
377       On Error Resume Next
378       If Not ParamRange Is Nothing Then DefaultValue = ParamRange.Address
379       Set NewRange = Application.InputBox(prompt:="Please select the 'parameter' cells that you will be changing between successsive solves of the model.", Type:=8, Default:=ParamRange.Address, Title:="OpenSolver Quick Solve Parameters")
382       On Error GoTo ErrorHandler
          
384       If Not NewRange Is Nothing Then
385           If NewRange.Worksheet.Name <> sheet.Name Then
386               Err.Raise OpenSolver_BuildError, Description:="The parameter cells need to be on the current worksheet."
388           End If
389           SetQuickSolveParameters NewRange, sheet

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
