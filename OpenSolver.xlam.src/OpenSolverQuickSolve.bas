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
378       If Not ParamRange Is Nothing Then DefaultValue = ParamRange.Address
          On Error Resume Next
379       Set NewRange = Application.InputBox(prompt:="Please select the 'parameter' cells that you will be changing between successsive solves of the model.", Type:=8, Default:=DefaultValue, Title:="OpenSolver Quick Solve Parameters")
          If Err.Number <> 0 And Err.Number <> 424 Then ' Error 424: Object required - happens on cancel press
              On Error GoTo ErrorHandler
              Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
          End If
          On Error GoTo ErrorHandler
          
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
