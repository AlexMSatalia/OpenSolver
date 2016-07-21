Attribute VB_Name = "SolverNOMAD"
Option Explicit

Public OS As COpenSolver
Dim IterationCount As Long

#If Mac Then
    Dim InteractiveStatus As Boolean
#End If

'NOMAD return status codes
Private Enum NomadResult
    LogFileError = -12  ' Error opening the log file, special code because we can't log the error to see what it is!
    UserCancelled = -3
    Optimal = 0
    ErrorOccured = 1
    SolveStoppedIter = 2         ' Limited by iterations
    SolveStoppedTime = 3         ' Limited by time
    Infeasible = 4               ' Provably infeasible
    SolveStoppedIterInfeas = 10  ' Infeasible after iter limit reached
    SolveStoppedTimeInfeas = 11  ' Infeasible after time limit reached
End Enum

Private Type VariableData
    X() As Variant
    NumVars As Long
End Type

' NOMAD functions
' Do not raise errors out of these functions, this will crash the NOMAD plugin
' We can raise errors inside the function and catch them, e.g. to detect escape presses

' Returns -1 on error, 0 otherwise
Function NOMAD_UpdateVar(X As Variant, Optional BestSolution As Variant = Nothing, Optional Infeasible As Boolean = False)
          On Error GoTo ErrorHandler
          Application.EnableCancelKey = xlErrorHandler
          
7115      IterationCount = IterationCount + 1

          ' Update solution
7116      Dim status As String
7117      status = "OpenSolver: Running NOMAD. Iteration " & IterationCount & "."
          ' Check for BestSolution = Nothing
7118      If Not VarType(BestSolution) = 9 Then
              ' Flip solution if maximisation
7119          If OS.ObjectiveSense = MaximiseObjective Then BestSolution = -BestSolution

7121          If Infeasible Then
7122              status = status & " Distance to feasibility: " & BestSolution
              Else
                  status = status & " Best solution so far: " & BestSolution
7123          End If
7124      End If
7125      UpdateStatusBar status, (IterationCount = 1)

          Dim i As Long, NumVars As Long
          'set new variable values on sheet
2452      NumVars = UBound(X)
2453      i = 1
          Dim AdjCell As Range
          ' If only one variable is returned, X is treated as a 1D array rather than 2D, so we need to access it
          ' differently.
2454      If NumVars = 1 Then
2455          For Each AdjCell In OS.AdjustableCells
2456              AdjCell.Value2 = X(i)
2457              i = i + 1
2458          Next AdjCell
2459      Else
2460          For Each AdjCell In OS.AdjustableCells
2461              AdjCell.Value2 = X(i, 1)
2462              i = i + 1
2463          Next AdjCell
2464      End If
          NOMAD_UpdateVar = 0

ExitFunction:
          Exit Function

ErrorHandler:
          If Not ReportError("SolverNOMAD", "NOMAD_UpdateVar") Then Resume
          NOMAD_UpdateVar = -1&
          Resume ExitFunction
End Function

' Returns -1 on error, array of new constraint values otherwise
Function NOMAD_GetValues() As Variant
          On Error GoTo ErrorHandler
          Application.EnableCancelKey = xlErrorHandler
          
          Dim X As Variant, i As Long, j As Long, k As Long, NumCons As Variant
2465      NumCons = NOMAD_GetNumConstraints()
2466      ReDim X(1 To NumCons(1, 1), 1 To 1)

          ' We get the objective value without validation and report an error value
          ' directly to NOMAD. Attempting to manipulate it can cause errors
          Dim ObjValue As Variant
          ObjValue = OS.GetCurrentObjectiveValue(ValidateNumeric:=False)
          
2469      If VarType(ObjValue) = vbDouble Then
2470          Select Case OS.ObjectiveSense
                  Case MaximiseObjective:
                      ' NOMAD only does minimization so flip sign
2472                  ObjValue = -ObjValue
                  Case TargetObjective:
                      ObjValue = Abs(ObjValue - OS.ObjectiveTargetValue)
              End Select
2481      End If
          X(1, 1) = ObjValue

2483      k = 1 'keep a count of what constraint its up to not including bounds
          Dim row As Long, constraint As Long
2484      row = 1
          Dim CurrentLHSValues As Variant
          Dim CurrentRHSValues As Variant
2485      For constraint = 1 To OS.NumConstraints
              ' Check to see what is different and add rows to sparsea
2486          If Not OS.LHSRange(constraint) Is Nothing Then ' skip Binary and Integer constraints
                  ' Get current value(s) for LHS and RHS of this constraint off the sheet. LHS is always an array (even if 1x1)
2487              OS.GetCurrentConstraintValues constraint, CurrentLHSValues, CurrentRHSValues, ValidateNumeric:=False
2488              If OS.LHSType(constraint) = SolverInputType.MultiCellRange Then
2489                  For i = 1 To UBound(CurrentLHSValues, 1)
2490                      For j = 1 To UBound(CurrentLHSValues, 2)
2491                          If OS.RowSetsBound(row) = False Then
2492                              If OS.RHSType(constraint) <> SolverInputType.MultiCellRange Then
2493                                  SetConstraintValue X, k, CurrentRHSValues, CurrentLHSValues(i, j), OS.Relation(constraint)
2494                              ElseIf UBound(CurrentLHSValues, 1) = UBound(CurrentRHSValues, 1) Then
2495                                  SetConstraintValue X, k, CurrentRHSValues(i, j), CurrentLHSValues(i, j), OS.Relation(constraint)
2496                              Else
2497                                  SetConstraintValueMismatchedDims X, k, CurrentRHSValues, CurrentLHSValues, OS.Relation(constraint), i, j
2498                              End If
2499                          End If
2500                          row = row + 1
2501                      Next j
2502                  Next i
2503              Else
2504                  If OS.RowSetsBound(row) = False Then
2505                      SetConstraintValue X, k, CurrentRHSValues, CurrentLHSValues(1, 1), OS.Relation(constraint)
2506                  End If
2507                  row = row + 1
2508              End If
2509          End If
2510      Next constraint
          
          'Get back new objective and difference between LHS and RHS values
2511      NOMAD_GetValues = X

ExitFunction:
          Exit Function
          
ErrorHandler:
          If Not ReportError("SolverNOMAD", "NOMAD_GetValues") Then Resume
          NOMAD_GetValues = -1&
          Resume ExitFunction
End Function

' Returns -1 on error, 0 otherwise
Function NOMAD_RecalculateValues()
          On Error GoTo ErrorHandler
          Application.EnableCancelKey = xlErrorHandler
          
          ForceCalculate "Warning: The worksheet calculation did not complete, and so the iteration may not be calculated correctly. Would you like to retry?"
          NOMAD_RecalculateValues = 0&
          
ExitFunction:
          Exit Function
          
ErrorHandler:
          If Not ReportError("SolverNOMAD", "NOMAD_RecalculateValues") Then Resume
          NOMAD_RecalculateValues = -1&
          Resume ExitFunction
End Function

' Returns -1 on error, integer number of variables otherwise
Function NOMAD_GetNumVariables() As Variant
          On Error GoTo ErrorHandler
          Application.EnableCancelKey = xlErrorHandler
          
          NOMAD_GetNumVariables = OS.AdjustableCells.Count

ExitFunction:
          Exit Function
          
ErrorHandler:
          If Not ReportError("SolverNOMAD", "NOMAD_GetNumVariables") Then Resume
          NOMAD_GetNumVariables = -1&
          Resume ExitFunction
End Function

' Returns -1 on error, integer number of constraints otherwise
Function NOMAD_GetNumConstraints() As Variant
          On Error GoTo ErrorHandler
          Application.EnableCancelKey = xlErrorHandler
          
          'Note: Bounds do not count as constraints and equalities count as 2 constraints
          Dim NumCons As Long
2540      NumCons = 0

          Dim row As Long, constraint As Long
2541      For row = 1 To OS.NumRows
              constraint = OS.RowToConstraint(row)
2542          If OS.RowSetsBound(row) = False Then NumCons = NumCons + 1
2543          If OS.Relation(constraint) = RelationEQ Then NumCons = NumCons + 1
2544      Next row

          'Number of objectives - NOMAD can do bi-objective
          'Note: Currently OpenSolver can only do single objectives- will need to set up multi objectives yourself
          Dim NumObjs As Long
2545      NumObjs = 1

          Dim X() As Variant
          ReDim X(1 To 1, 1 To 2)
          X(1, 1) = NumObjs + NumCons  ' Number of constraints and objectives
          X(1, 2) = NumObjs            ' Number of objectives
7131      NOMAD_GetNumConstraints = X

ExitFunction:
          Exit Function
          
ErrorHandler:
          If Not ReportError("SolverNOMAD", "NOMAD_GetNumConstraints") Then Resume
          NOMAD_GetNumConstraints = -1&
          Resume ExitFunction
End Function

' Returns -1 on error, array with variable data otherwise
Function NOMAD_GetVariableData() As Variant
          On Error GoTo ErrorHandler
          Application.EnableCancelKey = xlErrorHandler
          
          Dim data As VariableData
2547      data.NumVars = NOMAD_GetNumVariables

2548      ReDim data.X(1 To 4 * data.NumVars)

          Dim DefaultLowerBound As Double, DefaultUpperBound As Double
          DefaultLowerBound = IIf(OS.AssumeNonNegativeVars, 0, -10000000000000#)
          DefaultUpperBound = 10000000000000#

          Dim i As Long
          For i = 1 To OS.NumVars
              ' Set the default bounds
              SetLowerBound data, i, DefaultLowerBound
              SetUpperBound data, i, DefaultUpperBound
              
              Dim VarName As String
              VarName = OS.VarName(i)
              
              ' Set any specified bounds (overwriting defaults)
              If OS.VarLowerBounds.Exists(VarName) Then
                  SetLowerBound data, i, OS.VarLowerBounds.Item(VarName)
              End If
              If OS.VarUpperBounds.Exists(VarName) Then
                  SetUpperBound data, i, OS.VarUpperBounds.Item(VarName)
              End If
          
              ' Get the starting point
              If OS.InitialSolutionIsValid Then
                  SetStartingPoint data, i, OS.VarInitialValue(i)
              End If
          
              ' Set the variable type (int or bin)
2584          If OS.SolveRelaxation Then
                  ' Set Binary vars to have bounds of 0 and 1
                  If OS.VarCategory(i) = VarBinary Then
                      SetLowerBound data, i, 0
                      SetUpperBound data, i, 1
                  End If
                  SetVarType data, i, VarContinuous
              Else
                  SetVarType data, i, OS.VarCategory(i)
              End If
          Next i
          
2611      NOMAD_GetVariableData = data.X

ExitFunction:
          Exit Function
          
ErrorHandler:
          If Not ReportError("SolverNOMAD", "NOMAD_GetVariableData") Then Resume
          NOMAD_GetVariableData = -1&
          Resume ExitFunction
End Function

' Returns -1 on error, array of option data (string/length pairs for each key) otherwise
Function NOMAD_GetOptionData() As Variant
          On Error GoTo ErrorHandler
          Application.EnableCancelKey = xlErrorHandler
          
          IterationCount = 0
          
          ' On Mac we are re-entering the code so need to reset all our options
          #If Mac Then
              InteractiveStatus = Application.Interactive
              Application.Interactive = False
          #End If

          Dim SolverParameters As Dictionary
          Set SolverParameters = OS.SolverParameters
          ' Add extra values that depend on precision
          If SolverParameters.Exists(OS.Solver.PrecisionName) Then
              SolverParameters.Item("H_MIN") = SolverParameters.Item(OS.Solver.PrecisionName)
          End If
          
          Dim X() As Variant
          ReDim X(1 To SolverParameters.Count, 1 To 2)

          Dim i As Long, Key As Variant
          i = 1
          For Each Key In SolverParameters.Keys
              X(i, 1) = Key & " " & StrExNoPlus(SolverParameters.Item(Key))
              X(i, 2) = Len(X(i, 1))
              i = i + 1
          Next Key
          
          NOMAD_GetOptionData = X

ExitFunction:
          Exit Function
          
ErrorHandler:
          If Not ReportError("SolverNOMAD", "NOMAD_GetOptionData") Then Resume
          NOMAD_GetOptionData = -1&
          Resume ExitFunction
End Function

' Returns -1 on error, boolean whether to use warmstart or not otherwise
Function NOMAD_GetUseWarmstart() As Variant
          On Error GoTo ErrorHandler
          Application.EnableCancelKey = xlErrorHandler
          
          NOMAD_GetUseWarmstart = OS.InitialSolutionIsValid

ExitFunction:
          Exit Function
          
ErrorHandler:
          If Not ReportError("SolverNOMAD", "NOMAD_GetUseWarmstart") Then Resume
          NOMAD_GetUseWarmstart = -1&
          Resume ExitFunction
End Function

' Returns -1 if error, 0 otherwise
Function NOMAD_ShowCancelDialog()
          On Error GoTo ErrorHandler
          Application.EnableCancelKey = xlErrorHandler
          
          Dim Response As VbMsgBoxResult
          Response = ShowEscapeCancelMessage()
          If Response = vbYes Then
              OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError
              OpenSolverErrorHandler.ErrMsg = SILENT_ERROR
          End If
          NOMAD_ShowCancelDialog = 0&

ExitFunction:
          Exit Function
          
ErrorHandler:
          If Not ReportError("SolverNOMAD", "NOMAD_ShowCancelDialog") Then Resume
          NOMAD_ShowCancelDialog = -1&
          Resume ExitFunction
End Function

' Returns -1 if error, boolean indicating whether to abort otherwise
Function NOMAD_GetConfirmedAbort() As Variant
          On Error GoTo ErrorHandler
          Application.EnableCancelKey = xlErrorHandler
          
          ' Returns true if an abort has been confirmed by user
          NOMAD_GetConfirmedAbort = (OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError)

ExitFunction:
          Exit Function
          
ErrorHandler:
          If Not ReportError("SolverNOMAD", "NOMAD_GetConfirmedAbort") Then Resume
          NOMAD_GetConfirmedAbort = -1&
          Resume ExitFunction
End Function

' Returns -1 if error, array containing log path as string/length pair otherwise
Function NOMAD_GetLogFilePath() As Variant
          On Error GoTo ErrorHandler
          Application.EnableCancelKey = xlErrorHandler
          
          Dim X() As Variant
          ReDim X(1 To 1, 1 To 2)
          X(1, 1) = ConvertHfsPathToPosix(OS.LogFilePathName)
          X(1, 2) = Len(X(1, 1))
          NOMAD_GetLogFilePath = X

ExitFunction:
          Exit Function
          
ErrorHandler:
          If Not ReportError("SolverNOMAD", "NOMAD_GetLogFilePath") Then Resume
          NOMAD_GetLogFilePath = -1&
          Resume ExitFunction
End Function

Private Sub SetConstraintValue(ByRef ConstraintValues As Variant, ByRef k As Long, RHSValue As Variant, LHSValue As Variant, RelationType As Long)
                ' Sets the constraint value as appropriate for the given constraint (eg. LHS - RHS for <=) or returns
                ' "NaN" if either side contains an error (eg. #DIV/0!)
                ' This is for when the LHS and RHS ranges are the same dimension (both m x n)
2512            Select Case RelationType
                    Case RelationLE
2513                    ConstraintValues(k + 1, 1) = DifferenceOrError(LHSValue, RHSValue)
2514                Case RelationGE
2515                    ConstraintValues(k + 1, 1) = DifferenceOrError(RHSValue, LHSValue)
2516                Case RelationEQ
2517                    ConstraintValues(k + 1, 1) = DifferenceOrError(LHSValue, RHSValue)
2518                    ConstraintValues(k + 2, 1) = DifferenceOrError(RHSValue, LHSValue)
2519                    k = k + 1
2520            End Select
2521            k = k + 1
End Sub

Private Sub SetConstraintValueMismatchedDims(ByRef ConstraintValues As Variant, ByRef k As Long, RHSValues As Variant, LHSValues As Variant, RelationType As Long, i As Long, j As Long)
                ' Sets the constraint value as appropriate for the given constraint (eg. LHS - RHS for <=) or returns
                ' "NaN" if either side contains an error (eg. #DIV/0!)
                ' This is for when the LHS and RHS ranges have mismatched dimensions (m x n and n x m)
2522            Select Case RelationType
                    Case RelationLE
2523                    ConstraintValues(k + 1, 1) = DifferenceOrError(LHSValues(j, i), RHSValues(i, j))
2524                Case RelationGE
2525                    ConstraintValues(k + 1, 1) = DifferenceOrError(RHSValues(j, i), LHSValues(i, j))
2526                Case RelationEQ
2527                    ConstraintValues(k + 1, 1) = DifferenceOrError(LHSValues(j, i), RHSValues(i, j))
2528                    ConstraintValues(k + 2, 1) = DifferenceOrError(RHSValues(j, i), LHSValues(i, j))
2529                    k = k + 1
2530            End Select
2531            k = k + 1
End Sub

Private Function DifferenceOrError(Value1 As Variant, Value2 As Variant) As Variant
2532            If VarType(Value1) = vbError Then
2533                DifferenceOrError = Value1
2534            ElseIf VarType(Value2) = vbError Then
2535                DifferenceOrError = Value2
2536            Else
                    On Error GoTo ErrorHandler
2537                DifferenceOrError = Value1 - Value2
2538            End If
                Exit Function
                
ErrorHandler:
                DifferenceOrError = Value1
End Function

Private Function GetLowerBound(data As VariableData, i As Long) As Double
    GetLowerBound = data.X(i)
End Function

Private Sub SetLowerBound(ByRef data As VariableData, i As Long, value As Double)
    data.X(i) = value
End Sub

Private Function GetUpperBound(data As VariableData, i As Long) As Double
    GetUpperBound = data.X(data.NumVars + i)
End Function

Private Sub SetUpperBound(ByRef data As VariableData, i As Long, value As Double)
    data.X(data.NumVars + i) = value
End Sub

Private Function GetStartingPoint(data As VariableData, i As Long) As Double
    GetStartingPoint = data.X(2 * data.NumVars + i)
End Function

Private Sub SetStartingPoint(ByRef data As VariableData, i As Long, value As Double)
    data.X(2 * data.NumVars + i) = value
End Sub

Private Function GetVarType(data As VariableData, i As Long) As VariableType
    GetVarType = data.X(3 * data.NumVars + i)
End Function

Private Sub SetVarType(ByRef data As VariableData, i As Long, value As VariableType)
    data.X(3 * data.NumVars + i) = value
End Sub

#If Mac Then
Public Sub NOMAD_LoadResult(NomadRetVal As Long)
    On Error GoTo ErrorHandler

    Application.Interactive = InteractiveStatus
    GetNomadSolveResult NomadRetVal, OS
    OS.ReportAnySolutionSubOptimality

ExitSub:
    Set OS = Nothing
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    ReportError "OpenSolverAPI", "RunOpenSolver", True
    If OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError Then
        OS.SolveStatus = OpenSolverResult.AbortedThruUserAction
        Resume Next
    Else
        OS.SolveStatus = OpenSolverResult.ErrorOccurred
    End If
    Resume ExitSub
End Sub
#End If

Public Sub GetNomadSolveResult(NomadRetVal As Long, s As COpenSolver)
    'Catch any errors that occured while Nomad was solving
    Select Case NomadRetVal
    Case NomadResult.LogFileError
        s.SolveStatus = OpenSolverResult.ErrorOccurred
        Err.Raise OpenSolver_NomadError, Description:="NOMAD was unable to open the specified log file for writing: " & vbNewLine & vbNewLine & _
                                                      MakePathSafe(s.LogFilePathName)
    Case NomadResult.ErrorOccured
        s.SolveStatus = OpenSolverResult.ErrorOccurred
        
        ' Check logs for more info and raise an error if we find anything specific
        CheckLog s
        
        Err.Raise Number:=OpenSolver_NomadError, Description:="There was an error while Nomad was solving. No solution has been loaded into the sheet."
    Case NomadResult.SolveStoppedIter
        s.SolveStatus = OpenSolverResult.LimitedSubOptimal
        s.SolveStatusString = "NOMAD reached the maximum number of iterations and returned the best feasible solution it found. " & _
                              "This solution is not guaranteed to be an optimal solution." & vbNewLine & vbNewLine & _
                              "You can increase the maximum time and iterations under the options in the model dialogue or check whether your model is feasible."
        s.SolutionWasLoaded = True
    Case NomadResult.SolveStoppedTime
        s.SolveStatusString = "NOMAD reached the maximum time and returned the best feasible solution it found. " & _
                              "This solution is not guaranteed to be an optimal solution." & vbNewLine & vbNewLine & _
                              "You can increase the maximum time and iterations under the options in the model dialogue or check whether your model is feasible."
        s.SolveStatus = OpenSolverResult.LimitedSubOptimal
        s.SolutionWasLoaded = True
    Case NomadResult.Infeasible
        s.SolveStatusString = "Nomad could not find a feasible solution. " & _
                              "The best infeasible solution has been returned to the sheet." & vbNewLine & vbNewLine & _
                              "Try resolving at a different start point or check whether your model is feasible or relax some of your constraints."
        s.SolveStatus = OpenSolverResult.Infeasible
        s.SolutionWasLoaded = True
    Case NomadResult.SolveStoppedIterInfeas
        s.SolveStatusString = "Nomad reached the maximum number of iterations without finding a feasible solution. " & _
                              "The best infeasible solution has been returned to the sheet." & vbNewLine & vbNewLine & _
                              "You can increase the maximum number of iterations under the options in the model dialogue or check whether your model is feasible."
        s.SolveStatus = OpenSolverResult.Infeasible
        s.SolutionWasLoaded = True
    Case NomadResult.SolveStoppedTimeInfeas
        s.SolveStatusString = "Nomad reached the maximum time limit without finding a feasible solution. " & _
                              "The best infeasible solution has been returned to the sheet." & vbNewLine & vbNewLine & _
                              "You can increase the maximum time limit under the options in the model dialogue or check whether your model is feasible."
        s.SolveStatus = OpenSolverResult.Infeasible
        s.SolutionWasLoaded = True
    Case NomadResult.UserCancelled
        Err.Raise OpenSolver_UserCancelledError, "Running NOMAD", "Model solve cancelled by user."
    Case NomadResult.Optimal
        s.SolveStatus = OpenSolverResult.Optimal
        s.SolveStatusString = "Optimal"
    End Select
End Sub


Sub CheckLog(s As COpenSolver)
' If NOMAD encounters an error, we dump the exception to the log file. We can use this to deduce what went wrong
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler
    
    If Not FileOrDirExists(s.LogFilePathName) Then
        Err.Raise Number:=OpenSolver_SolveError, Description:="The solver did not create a log file. No new solution is available."
    End If
    
    Dim message As String
    Open s.LogFilePathName For Input As #3
        message = LCase(Input$(LOF(3), 3))
    Close #3
    
    If Not InStr(message, LCase("NOMAD")) > 0 Then GoTo ExitSub

    If InStr(message, LCase("invalid epsilon")) > 0 Then
        Err.Raise OpenSolver_NomadError, Description:="The specified precision was not a valid value for NOMAD. Check that you have specified a value above zero, or consult the NOMAD documentation for more information."
    End If
    
    If InStr(message, LCase("invalid parameter: DIMENSION")) > 0 Then
        Dim MaxSize As Long, Position As Long
        Position = InStrRev(message, " ")
        MaxSize = CLng(Mid(message, Position + 1, InStrRev(message, ")") - Position - 1))
        Err.Raise OpenSolver_NomadError, Description:="This model contains too many variables for NOMAD to solve. NOMAD is only capable of solving models with up to " & MaxSize & " variables."
    End If
    
    Dim Key As Variant
    For Each Key In s.SolverParameters.Keys()
        If InStr(message, LCase("invalid parameter: " & Key & " - unknown")) > 0 Then
            Err.Raise OpenSolver_NomadError, Description:="The parameter '" & Key & "' was not understood by NOMAD. Check that you have specified a valid parameter name, or consult the NOMAD documentation for more information."
        End If
        If InStr(message, LCase("invalid parameter: " & Key)) > 0 Then
            Err.Raise OpenSolver_NomadError, Description:="The value of the parameter '" & Key & "' supplied to NOMAD was invalid. Check that you have specified a valid value for this parameter, or consult the NOMAD documentation for more information."
        End If
    Next Key
        
    If InStr(message, LCase("invalid parameter")) > 0 Then
        Err.Raise OpenSolver_NomadError, Description:="One of the parameters supplied to NOMAD was invalid. This usually happens if the precision is too large. Try adjusting the values in the Solve Options dialog box."
    End If

ExitSub:
    Close #3
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Sub

ErrorHandler:
    If Not ReportError("CSolverNomad", "CheckLog") Then Resume
    RaiseError = True
    GoTo ExitSub
End Sub

