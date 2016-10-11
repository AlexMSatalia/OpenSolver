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
1         On Error GoTo ErrorHandler
2         Application.EnableCancelKey = xlErrorHandler
          
3         IterationCount = IterationCount + 1

          ' Update solution
4         Dim status As String
5         status = "OpenSolver: Running NOMAD. Iteration " & IterationCount & "."
          ' Check for BestSolution = Nothing
6         If Not VarType(BestSolution) = 9 Then
              ' Flip solution if maximisation
7             If OS.ObjectiveSense = MaximiseObjective Then BestSolution = -BestSolution

8             If Infeasible Then
9                 status = status & " Distance to feasibility: " & BestSolution
10            Else
11                status = status & " Best solution so far: " & BestSolution
12            End If
13        End If
14        UpdateStatusBar status, (IterationCount = 1)

          Dim i As Long, NumVars As Long
          'set new variable values on sheet
15        NumVars = UBound(X)
16        i = 1
          Dim AdjCell As Range
          ' If only one variable is returned, X is treated as a 1D array rather than 2D, so we need to access it
          ' differently.
17        If NumVars = 1 Then
18            For Each AdjCell In OS.AdjustableCells
19                AdjCell.Value2 = X(i)
20                i = i + 1
21            Next AdjCell
22        Else
23            For Each AdjCell In OS.AdjustableCells
24                AdjCell.Value2 = X(i, 1)
25                i = i + 1
26            Next AdjCell
27        End If
28        NOMAD_UpdateVar = 0

ExitFunction:
29        Exit Function

ErrorHandler:
30        If Not ReportError("SolverNOMAD", "NOMAD_UpdateVar") Then Resume
31        NOMAD_UpdateVar = -1&
32        Resume ExitFunction
End Function

' Returns -1 on error, array of new constraint values otherwise
Function NOMAD_GetValues() As Variant
1         On Error GoTo ErrorHandler
2         Application.EnableCancelKey = xlErrorHandler
          
          Dim X As Variant, i As Long, j As Long, k As Long, NumCons As Variant
3         NumCons = NOMAD_GetNumConstraints()
4         ReDim X(1 To NumCons(1, 1), 1 To 1)

          ' We get the objective value without validation and report an error value
          ' directly to NOMAD. Attempting to manipulate it can cause errors
          Dim ObjValue As Variant
5         ObjValue = OS.GetCurrentObjectiveValue(ValidateNumeric:=False)
          
6         If VarType(ObjValue) = vbDouble Then
7             Select Case OS.ObjectiveSense
                  Case MaximiseObjective:
                      ' NOMAD only does minimization so flip sign
8                     ObjValue = -ObjValue
9                 Case TargetObjective:
10                    ObjValue = Abs(ObjValue - OS.ObjectiveTargetValue)
11            End Select
12        End If
13        X(1, 1) = ObjValue

14        k = 1 'keep a count of what constraint its up to not including bounds
          Dim row As Long, constraint As Long
15        row = 1
          Dim CurrentLHSValues As Variant
          Dim CurrentRHSValues As Variant
16        For constraint = 1 To OS.NumConstraints
              ' Check to see what is different and add rows to sparsea
17            If Not OS.LHSRange(constraint) Is Nothing Then ' skip Binary and Integer constraints
                  ' Get current value(s) for LHS and RHS of this constraint off the sheet. LHS is always an array (even if 1x1)
18                OS.GetCurrentConstraintValues constraint, CurrentLHSValues, CurrentRHSValues, ValidateNumeric:=False
19                If OS.LHSType(constraint) = SolverInputType.MultiCellRange Then
20                    For i = 1 To UBound(CurrentLHSValues, 1)
21                        For j = 1 To UBound(CurrentLHSValues, 2)
22                            If OS.RowSetsBound(row) = False Then
23                                If OS.RHSType(constraint) <> SolverInputType.MultiCellRange Then
24                                    SetConstraintValue X, k, CurrentRHSValues, CurrentLHSValues(i, j), OS.Relation(constraint)
25                                ElseIf UBound(CurrentLHSValues, 1) = UBound(CurrentRHSValues, 1) Then
26                                    SetConstraintValue X, k, CurrentRHSValues(i, j), CurrentLHSValues(i, j), OS.Relation(constraint)
27                                Else
28                                    SetConstraintValueMismatchedDims X, k, CurrentRHSValues, CurrentLHSValues, OS.Relation(constraint), i, j
29                                End If
30                            End If
31                            row = row + 1
32                        Next j
33                    Next i
34                Else
35                    If OS.RowSetsBound(row) = False Then
36                        SetConstraintValue X, k, CurrentRHSValues, CurrentLHSValues(1, 1), OS.Relation(constraint)
37                    End If
38                    row = row + 1
39                End If
40            End If
41        Next constraint
          
          'Get back new objective and difference between LHS and RHS values
42        NOMAD_GetValues = X

ExitFunction:
43        Exit Function
          
ErrorHandler:
44        If Not ReportError("SolverNOMAD", "NOMAD_GetValues") Then Resume
45        NOMAD_GetValues = -1&
46        Resume ExitFunction
End Function

' Returns -1 on error, 0 otherwise
Function NOMAD_RecalculateValues()
1               On Error GoTo ErrorHandler
2               Application.EnableCancelKey = xlErrorHandler
                
3               ForceCalculate "Warning: The worksheet calculation did not complete, and so the iteration may not be calculated correctly. Would you like to retry?"
4               NOMAD_RecalculateValues = 0&
                
ExitFunction:
5               Exit Function
                
ErrorHandler:
6               If Not ReportError("SolverNOMAD", "NOMAD_RecalculateValues") Then Resume
7               NOMAD_RecalculateValues = -1&
8               Resume ExitFunction
End Function

' Returns -1 on error, integer number of variables otherwise
Function NOMAD_GetNumVariables() As Variant
1               On Error GoTo ErrorHandler
2               Application.EnableCancelKey = xlErrorHandler
                
3               NOMAD_GetNumVariables = OS.AdjustableCells.Count

ExitFunction:
4               Exit Function
                
ErrorHandler:
5               If Not ReportError("SolverNOMAD", "NOMAD_GetNumVariables") Then Resume
6               NOMAD_GetNumVariables = -1&
7               Resume ExitFunction
End Function

' Returns -1 on error, integer number of constraints otherwise
Function NOMAD_GetNumConstraints() As Variant
1         On Error GoTo ErrorHandler
2         Application.EnableCancelKey = xlErrorHandler
          
          'Note: Bounds do not count as constraints and equalities count as 2 constraints
          Dim NumCons As Long
3         NumCons = 0

          Dim row As Long, constraint As Long
4         For row = 1 To OS.NumRows
5             constraint = OS.RowToConstraint(row)
6             If OS.RowSetsBound(row) = False Then NumCons = NumCons + 1
7             If OS.Relation(constraint) = RelationEQ Then NumCons = NumCons + 1
8         Next row

          'Number of objectives - NOMAD can do bi-objective
          'Note: Currently OpenSolver can only do single objectives- will need to set up multi objectives yourself
          Dim NumObjs As Long
9         NumObjs = 1

          Dim X() As Variant
10        ReDim X(1 To 1, 1 To 2)
11        X(1, 1) = NumObjs + NumCons  ' Number of constraints and objectives
12        X(1, 2) = NumObjs            ' Number of objectives
13        NOMAD_GetNumConstraints = X

ExitFunction:
14        Exit Function
          
ErrorHandler:
15        If Not ReportError("SolverNOMAD", "NOMAD_GetNumConstraints") Then Resume
16        NOMAD_GetNumConstraints = -1&
17        Resume ExitFunction
End Function

' Returns -1 on error, array with variable data otherwise
Function NOMAD_GetVariableData() As Variant
1         On Error GoTo ErrorHandler
2         Application.EnableCancelKey = xlErrorHandler
          
          Dim data As VariableData
3         data.NumVars = NOMAD_GetNumVariables

4         ReDim data.X(1 To 4 * data.NumVars)

          Dim DefaultLowerBound As Double, DefaultUpperBound As Double
5         DefaultLowerBound = IIf(OS.AssumeNonNegativeVars, 0, -10000000000000#)
6         DefaultUpperBound = 10000000000000#

          Dim i As Long
7         For i = 1 To OS.NumVars
              ' Set the default bounds
8             SetLowerBound data, i, DefaultLowerBound
9             SetUpperBound data, i, DefaultUpperBound
              
              Dim VarName As String
10            VarName = OS.VarName(i)
              
              ' Set any specified bounds (overwriting defaults)
11            If OS.VarLowerBounds.Exists(VarName) Then
12                SetLowerBound data, i, OS.VarLowerBounds.Item(VarName)
13            End If
14            If OS.VarUpperBounds.Exists(VarName) Then
15                SetUpperBound data, i, OS.VarUpperBounds.Item(VarName)
16            End If
          
              ' Get the starting point
17            If OS.InitialSolutionIsValid Then
18                SetStartingPoint data, i, OS.VarInitialValue(i)
19            End If
          
              ' Set the variable type (int or bin)
20            If OS.SolveRelaxation Then
                  ' Set Binary vars to have bounds of 0 and 1
21                If OS.VarCategory(i) = VarBinary Then
22                    SetLowerBound data, i, 0
23                    SetUpperBound data, i, 1
24                End If
25                SetVarType data, i, VarContinuous
26            Else
27                SetVarType data, i, OS.VarCategory(i)
28            End If
29        Next i
          
30        NOMAD_GetVariableData = data.X

ExitFunction:
31        Exit Function
          
ErrorHandler:
32        If Not ReportError("SolverNOMAD", "NOMAD_GetVariableData") Then Resume
33        NOMAD_GetVariableData = -1&
34        Resume ExitFunction
End Function

' Returns -1 on error, array of option data (string/length pairs for each key) otherwise
Function NOMAD_GetOptionData() As Variant
1               On Error GoTo ErrorHandler
2               Application.EnableCancelKey = xlErrorHandler
                
3               IterationCount = 0
                
                ' On Mac we are re-entering the code so need to reset all our options
          #If Mac Then
4                   InteractiveStatus = Application.Interactive
5                   Application.Interactive = False
          #End If

                Dim SolverParameters As Dictionary
6               Set SolverParameters = OS.SolverParameters
                ' Add extra values that depend on precision
7               If SolverParameters.Exists(OS.Solver.PrecisionName) Then
8                   SolverParameters.Item("H_MIN") = SolverParameters.Item(OS.Solver.PrecisionName)
9               End If
                
                Dim X() As Variant
10              ReDim X(1 To SolverParameters.Count, 1 To 2)

                Dim i As Long, Key As Variant
11              i = 1
12              For Each Key In SolverParameters.Keys
13                  X(i, 1) = Key & " " & StrExNoPlus(SolverParameters.Item(Key))
14                  X(i, 2) = Len(X(i, 1))
15                  i = i + 1
16              Next Key
                
17              NOMAD_GetOptionData = X

ExitFunction:
18              Exit Function
                
ErrorHandler:
19              If Not ReportError("SolverNOMAD", "NOMAD_GetOptionData") Then Resume
20              NOMAD_GetOptionData = -1&
21              Resume ExitFunction
End Function

' Returns -1 on error, boolean whether to use warmstart or not otherwise
Function NOMAD_GetUseWarmstart() As Variant
1               On Error GoTo ErrorHandler
2               Application.EnableCancelKey = xlErrorHandler
                
3               NOMAD_GetUseWarmstart = OS.InitialSolutionIsValid

ExitFunction:
4               Exit Function
                
ErrorHandler:
5               If Not ReportError("SolverNOMAD", "NOMAD_GetUseWarmstart") Then Resume
6               NOMAD_GetUseWarmstart = -1&
7               Resume ExitFunction
End Function

' Returns -1 if error, 0 otherwise
Function NOMAD_ShowCancelDialog()
1               On Error GoTo ErrorHandler
2               Application.EnableCancelKey = xlErrorHandler
                
                Dim Response As VbMsgBoxResult
3               Response = ShowEscapeCancelMessage()
4               If Response = vbYes Then
5                   OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError
6                   OpenSolverErrorHandler.ErrMsg = SILENT_ERROR
7               End If
8               NOMAD_ShowCancelDialog = 0&

ExitFunction:
9               Exit Function
                
ErrorHandler:
10              If Not ReportError("SolverNOMAD", "NOMAD_ShowCancelDialog") Then Resume
11              NOMAD_ShowCancelDialog = -1&
12              Resume ExitFunction
End Function

' Returns -1 if error, boolean indicating whether to abort otherwise
Function NOMAD_GetConfirmedAbort() As Variant
1               On Error GoTo ErrorHandler
2               Application.EnableCancelKey = xlErrorHandler
                
                ' Returns true if an abort has been confirmed by user
3               NOMAD_GetConfirmedAbort = (OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError)

ExitFunction:
4               Exit Function
                
ErrorHandler:
5               If Not ReportError("SolverNOMAD", "NOMAD_GetConfirmedAbort") Then Resume
6               NOMAD_GetConfirmedAbort = -1&
7               Resume ExitFunction
End Function

' Returns -1 if error, array containing log path as string/length pair otherwise
Function NOMAD_GetLogFilePath() As Variant
1               On Error GoTo ErrorHandler
2               Application.EnableCancelKey = xlErrorHandler
                
                Dim X() As Variant
3               ReDim X(1 To 1, 1 To 2)
4               X(1, 1) = ConvertHfsPathToPosix(OS.LogFilePathName)
5               X(1, 2) = Len(X(1, 1))
6               NOMAD_GetLogFilePath = X

ExitFunction:
7               Exit Function
                
ErrorHandler:
8               If Not ReportError("SolverNOMAD", "NOMAD_GetLogFilePath") Then Resume
9               NOMAD_GetLogFilePath = -1&
10              Resume ExitFunction
End Function

Private Sub SetConstraintValue(ByRef ConstraintValues As Variant, ByRef k As Long, RHSValue As Variant, LHSValue As Variant, RelationType As Long)
                ' Sets the constraint value as appropriate for the given constraint (eg. LHS - RHS for <=) or returns
                ' "NaN" if either side contains an error (eg. #DIV/0!)
                ' This is for when the LHS and RHS ranges are the same dimension (both m x n)
1               Select Case RelationType
                    Case RelationLE
2                       ConstraintValues(k + 1, 1) = DifferenceOrError(LHSValue, RHSValue)
3                   Case RelationGE
4                       ConstraintValues(k + 1, 1) = DifferenceOrError(RHSValue, LHSValue)
5                   Case RelationEQ
6                       ConstraintValues(k + 1, 1) = DifferenceOrError(LHSValue, RHSValue)
7                       ConstraintValues(k + 2, 1) = DifferenceOrError(RHSValue, LHSValue)
8                       k = k + 1
9               End Select
10              k = k + 1
End Sub

Private Sub SetConstraintValueMismatchedDims(ByRef ConstraintValues As Variant, ByRef k As Long, RHSValues As Variant, LHSValues As Variant, RelationType As Long, i As Long, j As Long)
                ' Sets the constraint value as appropriate for the given constraint (eg. LHS - RHS for <=) or returns
                ' "NaN" if either side contains an error (eg. #DIV/0!)
                ' This is for when the LHS and RHS ranges have mismatched dimensions (m x n and n x m)
1               Select Case RelationType
                    Case RelationLE
2                       ConstraintValues(k + 1, 1) = DifferenceOrError(LHSValues(j, i), RHSValues(i, j))
3                   Case RelationGE
4                       ConstraintValues(k + 1, 1) = DifferenceOrError(RHSValues(j, i), LHSValues(i, j))
5                   Case RelationEQ
6                       ConstraintValues(k + 1, 1) = DifferenceOrError(LHSValues(j, i), RHSValues(i, j))
7                       ConstraintValues(k + 2, 1) = DifferenceOrError(RHSValues(j, i), LHSValues(i, j))
8                       k = k + 1
9               End Select
10              k = k + 1
End Sub

Private Function DifferenceOrError(Value1 As Variant, Value2 As Variant) As Variant
1               If VarType(Value1) = vbError Then
2                   DifferenceOrError = Value1
3               ElseIf VarType(Value2) = vbError Then
4                   DifferenceOrError = Value2
5               Else
6                   On Error GoTo ErrorHandler
7                   DifferenceOrError = Value1 - Value2
8               End If
9               Exit Function
                
ErrorHandler:
10              DifferenceOrError = Value1
End Function

Private Function GetLowerBound(data As VariableData, i As Long) As Double
1         GetLowerBound = data.X(i)
End Function

Private Sub SetLowerBound(ByRef data As VariableData, i As Long, value As Double)
1         data.X(i) = value
End Sub

Private Function GetUpperBound(data As VariableData, i As Long) As Double
1         GetUpperBound = data.X(data.NumVars + i)
End Function

Private Sub SetUpperBound(ByRef data As VariableData, i As Long, value As Double)
1         data.X(data.NumVars + i) = value
End Sub

Private Function GetStartingPoint(data As VariableData, i As Long) As Double
1         GetStartingPoint = data.X(2 * data.NumVars + i)
End Function

Private Sub SetStartingPoint(ByRef data As VariableData, i As Long, value As Double)
1         data.X(2 * data.NumVars + i) = value
End Sub

Private Function GetVarType(data As VariableData, i As Long) As VariableType
1         GetVarType = data.X(3 * data.NumVars + i)
End Function

Private Sub SetVarType(ByRef data As VariableData, i As Long, value As VariableType)
1         data.X(3 * data.NumVars + i) = value
End Sub

#If Mac Then
Public Sub NOMAD_LoadResult(NomadRetVal As Long)
1         On Error GoTo ErrorHandler

2         Application.Interactive = InteractiveStatus
3         GetNomadSolveResult NomadRetVal, OS
4         OS.ReportAnySolutionSubOptimality

ExitSub:
5         Set OS = Nothing
6         Application.StatusBar = False
7         Exit Sub
          
ErrorHandler:
8         ReportError "OpenSolverAPI", "RunOpenSolver", True
9         If OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError Then
10            OS.SolveStatus = OpenSolverResult.AbortedThruUserAction
11            Resume Next
12        Else
13            OS.SolveStatus = OpenSolverResult.ErrorOccurred
14        End If
15        Resume ExitSub
End Sub
#End If

Public Sub GetNomadSolveResult(NomadRetVal As Long, s As COpenSolver)
          'Catch any errors that occured while Nomad was solving
1         Select Case NomadRetVal
          Case NomadResult.LogFileError
2             s.SolveStatus = OpenSolverResult.ErrorOccurred
3             RaiseGeneralError "NOMAD was unable to open the specified log file for writing: " & vbNewLine & vbNewLine & _
                                MakePathSafe(s.LogFilePathName)
4         Case NomadResult.ErrorOccured
5             s.SolveStatus = OpenSolverResult.ErrorOccurred
              
              ' Check logs for more info and raise an error if we find anything specific
6             CheckLog s
              
7             RaiseGeneralError "There was an error while Nomad was solving. No solution has been loaded into the sheet."
8         Case NomadResult.SolveStoppedIter
9             s.SolveStatus = OpenSolverResult.LimitedSubOptimal
10            s.SolveStatusString = "NOMAD reached the maximum number of iterations and returned the best feasible solution it found. " & _
                                    "This solution is not guaranteed to be an optimal solution." & vbNewLine & vbNewLine & _
                                    "You can increase the maximum time and iterations under the options in the model dialogue or check whether your model is feasible."
11            s.SolutionWasLoaded = True
12        Case NomadResult.SolveStoppedTime
13            s.SolveStatusString = "NOMAD reached the maximum time and returned the best feasible solution it found. " & _
                                    "This solution is not guaranteed to be an optimal solution." & vbNewLine & vbNewLine & _
                                    "You can increase the maximum time and iterations under the options in the model dialogue or check whether your model is feasible."
14            s.SolveStatus = OpenSolverResult.LimitedSubOptimal
15            s.SolutionWasLoaded = True
16        Case NomadResult.Infeasible
17            s.SolveStatusString = "Nomad could not find a feasible solution. " & _
                                    "The best infeasible solution has been returned to the sheet." & vbNewLine & vbNewLine & _
                                    "Try resolving at a different start point or check whether your model is feasible or relax some of your constraints."
18            s.SolveStatus = OpenSolverResult.Infeasible
19            s.SolutionWasLoaded = True
20        Case NomadResult.SolveStoppedIterInfeas
21            s.SolveStatusString = "Nomad reached the maximum number of iterations without finding a feasible solution. " & _
                                    "The best infeasible solution has been returned to the sheet." & vbNewLine & vbNewLine & _
                                    "You can increase the maximum number of iterations under the options in the model dialogue or check whether your model is feasible."
22            s.SolveStatus = OpenSolverResult.Infeasible
23            s.SolutionWasLoaded = True
24        Case NomadResult.SolveStoppedTimeInfeas
25            s.SolveStatusString = "Nomad reached the maximum time limit without finding a feasible solution. " & _
                                    "The best infeasible solution has been returned to the sheet." & vbNewLine & vbNewLine & _
                                    "You can increase the maximum time limit under the options in the model dialogue or check whether your model is feasible."
26            s.SolveStatus = OpenSolverResult.Infeasible
27            s.SolutionWasLoaded = True
28        Case NomadResult.UserCancelled
29            RaiseUserCancelledError
30        Case NomadResult.Optimal
31            s.SolveStatus = OpenSolverResult.Optimal
32            s.SolveStatusString = "Optimal"
33        End Select
End Sub


Sub CheckLog(s As COpenSolver)
      ' If NOMAD encounters an error, we dump the exception to the log file. We can use this to deduce what went wrong
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
3         If Not FileOrDirExists(s.LogFilePathName) Then
4             RaiseGeneralError "The solver did not create a log file. No new solution is available.", _
                                LINK_NO_SOLUTION_FILE
5         End If
          
          Dim message As String
6         Open s.LogFilePathName For Input As #3
7             message = LCase(Input$(LOF(3), 3))
8         Close #3
          
9         If Not InStr(message, LCase("NOMAD")) > 0 Then GoTo ExitSub

10        If InStr(message, LCase("invalid epsilon")) > 0 Then
11            RaiseUserError "The specified precision was not a valid value for NOMAD. Check that you have specified a value above zero, or consult the NOMAD documentation for more information.", _
                              LINK_PARAMETER_DOCS
12        End If
          
13        If InStr(message, LCase("invalid parameter: DIMENSION")) > 0 Then
              Dim MaxSize As Long, Position As Long
14            Position = InStrRev(message, " ")
15            MaxSize = CLng(Mid(message, Position + 1, InStrRev(message, ")") - Position - 1))
16            RaiseUserError "This model contains too many variables for NOMAD to solve. NOMAD is only capable of solving models with up to " & MaxSize & " variables."
17        End If
          
          Dim Key As Variant
18        For Each Key In s.SolverParameters.Keys()
19            If InStr(message, LCase("invalid parameter: " & Key & " - unknown")) > 0 Then
20                RaiseUserError "The parameter '" & Key & "' was not understood by NOMAD. Check that you have specified a valid parameter name, or consult the NOMAD documentation for more information.", _
                                 LINK_PARAMETER_DOCS
21            End If
22            If InStr(message, LCase("invalid parameter: " & Key)) > 0 Then
23                RaiseUserError "The value of the parameter '" & Key & "' supplied to NOMAD was invalid. Check that you have specified a valid value for this parameter, or consult the NOMAD documentation for more information.", _
                                 LINK_PARAMETER_DOCS
24            End If
25        Next Key
              
26        If InStr(message, LCase("invalid parameter")) > 0 Then
27            RaiseUserError "One of the parameters supplied to NOMAD was invalid. This usually happens if the precision is too large. Try adjusting the values in the Solve Options dialog box."
28        End If

ExitSub:
29        Close #3
30        If RaiseError Then RethrowError
31        Exit Sub

ErrorHandler:
32        If Not ReportError("CSolverNomad", "CheckLog") Then Resume
33        RaiseError = True
34        GoTo ExitSub
End Sub

