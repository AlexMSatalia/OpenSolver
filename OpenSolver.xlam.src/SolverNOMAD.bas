Attribute VB_Name = "SolverNOMAD"
Option Explicit

Public OS As COpenSolver
Dim IterationCount As Long

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
    numVars As Long
End Type

' NOMAD functions
' Do not put error handling in these functions, there is already error-handling in the NOMAD DLL

Sub NOMAD_UpdateVar(X As Variant, Optional BestSolution As Variant = Nothing, Optional Infeasible As Boolean = False)
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

          Dim i As Long, numVars As Long
          'set new variable values on sheet
2452      numVars = UBound(X)
2453      i = 1
          Dim AdjCell As Range
          ' If only one variable is returned, X is treated as a 1D array rather than 2D, so we need to access it
          ' differently.
2454      If numVars = 1 Then
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
End Sub

Function NOMAD_GetValues() As Variant
          Dim X As Variant, i As Long, j As Long, k As Long, numCons As Variant
2465      numCons = NOMAD_GetNumConstraints()
2466      ReDim X(1 To numCons(0), 1 To 1)
          '====NOMAD only does minimise so need to change objective if it is max====
          ' If no objective, just set a constant.
          ' TODO: fix this to set it based on amount of violation to hunt for feasibility
2467      If OS.ObjRange Is Nothing Then
2468          X(1, 1) = 0
          ' If objective cell is error, report this directly to NOMAD. Attempting to manipulate it can cause errors
2469      ElseIf VarType(OS.ObjRange.Value2) = vbError Then
2470          X(1, 1) = OS.ObjRange.Value2
          'If objective sense is maximise then multiply by minus 1
2471      ElseIf OS.ObjectiveSense = MaximiseObjective Then
2472          If OS.ObjRange.Value2 <> 0 Then
2473              X(1, 1) = -1 * OS.ObjRange.Value2 'objective value
2474          Else
2475              X(1, 1) = OS.ObjRange.Value2
2476          End If
          'Else if objective sense is minimise leave it
2477      ElseIf OS.ObjectiveSense = MinimiseObjective Then
2478          X(1, 1) = OS.ObjRange.Value2
2479      ElseIf OS.ObjectiveSense = TargetObjective Then
2480          X(1, 1) = Abs(OS.ObjRange.Value2 - OS.ObjectiveTargetValue)
2481      End If
2483      k = 1 'keep a count of what constraint its up to not including bounds
          Dim row As Long, constraint As Long
2484      row = 1
          Dim CurrentLHSValues As Variant
          Dim CurrentRHSValues As Variant
2485      For constraint = 1 To OS.NumConstraints
              ' Check to see what is different and add rows to sparsea
2486          If Not OS.LHSRange(constraint) Is Nothing Then ' skip Binary and Integer constraints
                  ' Get current value(s) for LHS and RHS of this constraint off the sheet. LHS is always an array (even if 1x1)
2487              OS.GetCurrentConstraintValues constraint, CurrentLHSValues, CurrentRHSValues
2488              If OS.LHSType(constraint) = SolverInputType.MultiCellRange Then
2489                  For i = 1 To UBound(OS.LHSOriginalValues(constraint), 1)
2490                      For j = 1 To UBound(OS.LHSOriginalValues(constraint), 2)
2491                          If OS.RowSetsBound(row) = False Then
2492                              If OS.RHSType(constraint) <> SolverInputType.MultiCellRange Then
2493                                  SetConstraintValue X, k, CurrentRHSValues, CurrentLHSValues(i, j), OS.Relation(row)
2494                              ElseIf UBound(OS.LHSOriginalValues(constraint), 1) = UBound(OS.RHSOriginalValues(constraint), 1) Then
2495                                  SetConstraintValue X, k, CurrentRHSValues(i, j), CurrentLHSValues(i, j), OS.Relation(row)
2496                              Else
2497                                  SetConstraintValueMismatchedDims X, k, CurrentRHSValues, CurrentLHSValues, OS.Relation(row), i, j
2498                              End If
2499                          End If
2500                          row = row + 1
2501                      Next j
2502                  Next i
2503              Else
2504                  If OS.RowSetsBound(row) = False Then
2505                      SetConstraintValue X, k, CurrentRHSValues, CurrentLHSValues(1, 1), OS.Relation(row)
2506                  End If
2507                  row = row + 1
2508              End If
2509          End If
2510      Next constraint
          
          'Get back new objective and difference between LHS and RHS values
2511      NOMAD_GetValues = X
End Function

Sub NOMAD_RecalculateValues()
          On Error Resume Next  ' Eat any errors - we can't throw an error while the C++ code is running
7129      If Not ForceCalculate("Warning: The worksheet calculation did not complete, and so the iteration may not be calculated correctly. Would you like to retry?") Then Exit Sub
          On Error GoTo 0
End Sub

Function NOMAD_GetNumVariables() As Variant
7130      NOMAD_GetNumVariables = OS.AdjustableCells.Count
End Function

Function NOMAD_GetNumConstraints() As Variant
          'The number of constraints is actually the number of Objectives + Number of Constraints
          'Note: Bounds do not count as constraints and equalities count as 2 constraints
          Dim row As Long
          Dim X(0 To 1) As Double
2540      X(0) = 1
2541      For row = 1 To OS.NumRows
2542          If OS.RowSetsBound(row) = False Then X(0) = X(0) + 1
2543          If OS.Relation(row) = RelationEQ Then X(0) = X(0) + 1
2544      Next row

          'Number of objectives - NOMAD can do bi-objective
          'Note: Currently OpenSolver can only do single objectives- will need to set up multi objectives yourself
2545      X(1) = 1 'number of objectives

7131      NOMAD_GetNumConstraints = X
End Function

Function NOMAD_GetVariableData() As Variant
          Dim data As VariableData
2547      data.numVars = NOMAD_GetNumVariables

2548      ReDim data.X(1 To 4 * data.numVars)

          Dim DefaultLowerBound As Double, DefaultUpperBound As Double
          DefaultLowerBound = IIf(OS.AssumeNonNegativeVars, 0, -10000000000000#)
          DefaultUpperBound = 10000000000000#

          Dim i As Long, c As Range
          i = 1
          For Each c In OS.AdjustableCells
              ' Set the default bounds
              SetLowerBound data, i, DefaultLowerBound
              SetUpperBound data, i, DefaultUpperBound
              
              ' Set any specified bounds (overwriting defaults)
              If TestKeyExists(OS.VarLowerBounds, c.Address) Then
                  SetLowerBound data, i, OS.VarLowerBounds(c.Address)
              End If
              If TestKeyExists(OS.VarUpperBounds, c.Address) Then
                  SetUpperBound data, i, OS.VarUpperBounds(c.Address)
              End If
              
              Dim LowerBound As Double, UpperBound As Double
              LowerBound = GetLowerBound(data, i)
              UpperBound = GetUpperBound(data, i)
          
              ' Get the starting point
              SetStartingPoint data, i, c.value
          
              ' Initialise all variables as continuous
              SetVarType data, i, VariableType.VarContinuous
              
              ' Get the variable type (int or bin)
2584          If OS.SolveRelaxation Then
                  ' Set Binary vars to have bounds of 0 and 1 and start at 0
                  If TestIntersect(c, OS.BinaryCellsRange) Then
                      SetLowerBound data, i, 0
                      SetUpperBound data, i, 1
                      SetStartingPoint data, i, 0
                  End If
              Else
                  Dim Integral As Boolean
                  Integral = True
                  If TestIntersect(c, OS.BinaryCellsRange) Then
                      SetVarType data, i, VariableType.VarBinary
                  ElseIf TestIntersect(c, OS.IntegerCellsRange) Then
                      SetVarType data, i, VariableType.VarInteger
                  Else
                      Integral = False
                  End If
                  
                  If Integral Then
                      'Make bounds on integer and binary constraints integer
2587                  If LowerBound > 0 Then
2588                      LowerBound = Application.WorksheetFunction.RoundUp(LowerBound, 0)
2589                  Else
2590                      LowerBound = Application.WorksheetFunction.RoundDown(LowerBound, 0)
2591                  End If
                      SetLowerBound data, i, LowerBound

2592                  If UpperBound > 0 Then
2593                      UpperBound = Application.WorksheetFunction.RoundDown(UpperBound, 0)
2594                  Else
2595                      UpperBound = Application.WorksheetFunction.RoundUp(UpperBound, 0)
2596                  End If
                      SetUpperBound data, i, UpperBound
                      
                      'Make starting positions on integer and binary constraints integer
                      SetStartingPoint data, i, Round(GetStartingPoint(data, i))
                  End If
              End If
              
              ' Force starting point between the bounds
              Dim StartingPoint As Double
              StartingPoint = GetStartingPoint(data, i)
              If StartingPoint < LowerBound Then
                  StartingPoint = LowerBound
              ElseIf StartingPoint > UpperBound Then
                  StartingPoint = UpperBound
              End If
              SetStartingPoint data, i, StartingPoint
              
              i = i + 1
          Next c
          
2611      NOMAD_GetVariableData = data.X
End Function

Function NOMAD_GetOptionData() As Variant
          IterationCount = 0

          Dim SolverParameters As Dictionary
          Set SolverParameters = OS.SolverParameters
          ' Add extra values that depend on precision
          If SolverParameters.Exists(OS.Solver.PrecisionName) Then
              SolverParameters.Add Key:="H_MIN", Item:=SolverParameters.Item(OS.Solver.PrecisionName)
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
End Function

Function NOMAD_ShowCancelDialog() As Variant
          Dim Response As VbMsgBoxResult
          Response = ShowEscapeCancelMessage()
          NOMAD_ShowCancelDialog = (Response = vbYes)
End Function

Function NOMAD_GetLogFilePath() As Variant
          Dim X() As Variant
          ReDim X(1 To 1, 1 To 2)
          X(1, 1) = ConvertHfsPathToPosix(OS.LogFilePathName)
          X(1, 2) = Len(X(1, 1))
          NOMAD_GetLogFilePath = X
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
    GetUpperBound = data.X(data.numVars + i)
End Function

Private Sub SetUpperBound(ByRef data As VariableData, i As Long, value As Double)
    data.X(data.numVars + i) = value
End Sub

Private Function GetStartingPoint(data As VariableData, i As Long) As Double
    GetStartingPoint = data.X(2 * data.numVars + i)
End Function

Private Sub SetStartingPoint(ByRef data As VariableData, i As Long, value As Double)
    data.X(2 * data.numVars + i) = value
End Sub

Private Function GetVarType(data As VariableData, i As Long) As Double
    GetVarType = data.X(3 * data.numVars + i)
End Function

Private Sub SetVarType(ByRef data As VariableData, i As Long, value As Double)
    data.X(3 * data.numVars + i) = value
End Sub

Public Sub NOMAD_LoadResult(NomadRetVal As Long)
    'Catch any errors that occured while Nomad was solving
    Select Case NomadRetVal
    Case NomadResult.LogFileError
        OS.SolveStatus = OpenSolverResult.ErrorOccurred
        Err.Raise OpenSolver_NomadError, Description:="NOMAD was unable to open the specified log file for writing: " & vbNewLine & vbNewLine & _
                                                      OS.LogFilePathName
    Case NomadResult.ErrorOccured
        OS.SolveStatus = OpenSolverResult.ErrorOccurred
        
        ' Check logs for more info and raise an error if we find anything specific
        CheckLog OS
        
        Err.Raise Number:=OpenSolver_NomadError, Description:="There was an error while Nomad was solving. No solution has been loaded into the sheet."
    Case NomadResult.SolveStoppedIter
        OS.SolveStatus = OpenSolverResult.LimitedSubOptimal
        OS.SolveStatusString = "NOMAD reached the maximum number of iterations and returned the best feasible solution it found. " & _
                              "This solution is not guaranteed to be an optimal solution." & vbNewLine & vbNewLine & _
                              "You can increase the maximum time and iterations under the options in the model dialogue or check whether your model is feasible."
        OS.SolutionWasLoaded = True
    Case NomadResult.SolveStoppedTime
        OS.SolveStatusString = "NOMAD reached the maximum time and returned the best feasible solution it found. " & _
                              "This solution is not guaranteed to be an optimal solution." & vbNewLine & vbNewLine & _
                              "You can increase the maximum time and iterations under the options in the model dialogue or check whether your model is feasible."
        OS.SolveStatus = OpenSolverResult.LimitedSubOptimal
        OS.SolutionWasLoaded = True
    Case NomadResult.Infeasible
        OS.SolveStatusString = "Nomad could not find a feasible solution. " & _
                              "The best infeasible solution has been returned to the sheet." & vbNewLine & vbNewLine & _
                              "Try resolving at a different start point or check whether your model is feasible or relax some of your constraints."
        OS.SolveStatus = OpenSolverResult.Infeasible
        OS.SolutionWasLoaded = True
    Case NomadResult.SolveStoppedIterInfeas
        OS.SolveStatusString = "Nomad reached the maximum number of iterations without finding a feasible solution. " & _
                              "The best infeasible solution has been returned to the sheet." & vbNewLine & vbNewLine & _
                              "You can increase the maximum number of iterations under the options in the model dialogue or check whether your model is feasible."
        OS.SolveStatus = OpenSolverResult.Infeasible
        OS.SolutionWasLoaded = True
    Case NomadResult.SolveStoppedTimeInfeas
        OS.SolveStatusString = "Nomad reached the maximum time limit without finding a feasible solution. " & _
                              "The best infeasible solution has been returned to the sheet." & vbNewLine & vbNewLine & _
                              "You can increase the maximum time limit under the options in the model dialogue or check whether your model is feasible."
        OS.SolveStatus = OpenSolverResult.Infeasible
        OS.SolutionWasLoaded = True
    Case NomadResult.UserCancelled
        Err.Raise OpenSolver_UserCancelledError, "Running NOMAD", "Model solve cancelled by user."
    Case NomadResult.Optimal
        OS.SolveStatus = OpenSolverResult.Optimal
        OS.SolveStatusString = "Optimal"
    End Select

    #If Mac Then
        OS.ReportAnySolutionSubOptimality
        Set OS = Nothing
        Application.StatusBar = False
    #End If
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
        message = Input$(LOF(3), 3)
    Close #3
    
    If Not InStrText(message, "NOMAD") > 0 Then GoTo ExitSub

    If InStrText(message, "invalid parameter: DIMENSION") Then
        Dim MaxSize As Long, Position As Long
        Position = InStrRev(message, " ")
        MaxSize = CInt(Mid(message, Position + 1, InStrRev(message, ")") - Position - 1))
        Err.Raise OpenSolver_NomadError, Description:="This model contains too many variables for NOMAD to solve. NOMAD is only capable of solving models with up to " & MaxSize & " variables."
    End If
    
    Dim Key As Variant
    For Each Key In s.SolverParameters.Keys()
        If InStrText(message, "invalid parameter: " & UCase(Key) & " - unknown") Then
            Err.Raise OpenSolver_NomadError, Description:="The parameter '" & UCase(Key) & "' was not understood by NOMAD. Check that you have specified a valid parameter name, or consult the NOMAD documentation for more information."
        End If
        If InStrText(message, "invalid parameter: " & UCase(Key)) Then
            Err.Raise OpenSolver_NomadError, Description:="The value of the parameter '" & UCase(Key) & "' supplied to NOMAD was invalid. Check that you have specified a valid value for this parameter, or consult the NOMAD documentation for more information."
        End If
    Next Key
        
    If InStrText(message, "invalid parameter") Then
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

