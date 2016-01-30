Attribute VB_Name = "OpenSolverQuickSolve"
Option Explicit

Public QuickSolver As COpenSolver

Function SetQuickSolveParameterRange() As Boolean
          On Error GoTo ErrorHandler

362       SetQuickSolveParameterRange = False
          
          Dim sheet As Worksheet
          GetActiveSheetIfMissing sheet
          
          ' Find the Parameter range
          Dim ParamRangeRefersTo As String
375       ParamRangeRefersTo = GetQuickSolveParametersRefersTo(sheet)
          
          ' Get a range from the user
          Dim NewValue As String
          On Error Resume Next
379       NewValue = Application.InputBox( _
                         prompt:="Please select the 'parameter' cells that you will be changing between successsive solves of the model.", _
                         Type:=0, _
                         Default:=GetDisplayAddress(ParamRangeRefersTo, sheet, False), _
                         Title:="OpenSolver Quick Solve Parameters")
          If Err.Number <> 0 And Err.Number <> 424 Then ' Error 424: Object required - happens on cancel press
              On Error GoTo ErrorHandler
              Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
          End If
          On Error GoTo ErrorHandler
          
          ' Formula is always returned as ="<input>" or =<input> depending on whether the user
          ' entered an equals in the formula
          ' We make sure it's in A1 notation and strip the equals and any quotes
          NewValue = Application.ConvertFormula(NewValue, xlR1C1, xlA1)
          NewValue = Mid(NewValue, 2)
          If Left(NewValue, 1) = """" Then NewValue = Mid(NewValue, 2, Len(NewValue) - 2)
          
          SetQuickSolveParametersRefersTo RefEditToRefersTo(NewValue), sheet

          ' Return true as we have succeeded
393       SetQuickSolveParameterRange = True

ExitFunction:
          Exit Function

ErrorHandler:
          ReportError "OpenSolverQuickSolve", "SetQuickSolveParameterRange", True, False
          GoTo ExitFunction
End Function
