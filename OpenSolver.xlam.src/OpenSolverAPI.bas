Attribute VB_Name = "OpenSolverAPI"
Option Explicit

Public Const sOpenSolverVersion As String = "2.9.0"
Public Const sOpenSolverDate As String = "2017.11.10"

'/**
' * Solves the OpenSolver model on the current sheet.
' * @param {} SolveRelaxation If True, all integer and boolean constraints will be relaxed to allow continuous values for these variables. Defaults to False
' * @param {} MinimiseUserInteraction If True, all dialogs and messages will be suppressed. Use this when automating a lot of solves so that there are no interruptions. Defaults to False
' * @param {} LinearityCheckOffset Sets the base value used for checking if the model is linear. Change this if a non-linear model is not being detected as non-linear. Defaults to 10.423 (a random number that hopefully does not occur in the model, e.g. =ABS(A1-10.423))
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function RunOpenSolver(Optional SolveRelaxation As Boolean = False, _
                              Optional MinimiseUserInteraction As Boolean = False, _
                              Optional LinearityOffset As Double = 10.423, _
                              Optional sheet As Worksheet) As OpenSolverResult
1         CheckLocationValid  ' Check for unicode in path
          
2         ClearError
3         On Error GoTo ErrorHandler
          
          Dim InteractiveStatus As Boolean
4         InteractiveStatus = Application.Interactive
5         Application.Interactive = False
          
6         GetActiveSheetIfMissing sheet

          Dim CurrentSolver As String
          CurrentSolver = GetChosenSolver(sheet)
            
          If CurrentSolver = "NeosCplex" Then
              If Not ValidateEmail Then RaiseUserError "CPLEX requires an email address. Please check a valid email address has been input under Model > Options."
          End If

7         RunOpenSolver = OpenSolverResult.Unsolved
          Dim OpenSolver As COpenSolver
8         Set OpenSolver = New COpenSolver

9         OpenSolver.BuildModelFromSolverData LinearityOffset, GetLinearityCheck(sheet), MinimiseUserInteraction, SolveRelaxation, sheet
          ' Only proceed with solve if nothing detected while building model
10        If OpenSolver.SolveStatus = OpenSolverResult.Unsolved Then
11            SolveModel OpenSolver, SolveRelaxation, MinimiseUserInteraction
12        End If
          
13        RunOpenSolver = OpenSolver.SolveStatus
14        If Not MinimiseUserInteraction Then OpenSolver.ReportAnySolutionSubOptimality

ExitFunction:
15        Application.Interactive = InteractiveStatus
16        Set OpenSolver = Nothing    ' Free any OpenSolver memory used
17        Exit Function

ErrorHandler:
18        ReportError "OpenSolverAPI", "RunOpenSolver", True, MinimiseUserInteraction
19        If OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError Then
20            RunOpenSolver = AbortedThruUserAction
21        Else
22            RunOpenSolver = OpenSolverResult.ErrorOccurred
23        End If
24        GoTo ExitFunction
End Function

'/**
' * Gets a list of short names for all solvers that can be set
' */
Public Function GetAvailableSolvers() As String()
1         GetAvailableSolvers = StringArray("CBC", "Gurobi", "NeosCBC", "NeosCplex", "Bonmin", "Couenne", "NOMAD", "NeosBon", "NeosCou", "SolveEngine")
End Function

'/**
' * Gets the short name of the currently selected solver for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetChosenSolver(Optional sheet As Worksheet) As String
1         GetActiveSheetIfMissing sheet
2         If Not GetNamedStringIfExists(sheet, "OpenSolver_ChosenSolver", GetChosenSolver) Then GoTo SetDefault
          
          ' Check solver is an allowed solver
3         On Error GoTo SetDefault
4         WorksheetFunction.Match GetChosenSolver, GetAvailableSolvers, 0
5         Exit Function
          
SetDefault:
          ' See if we can choose based on the selection for Solver
          Dim SolverEng As Long
6         If GetNamedIntegerIfExists(sheet, "solver_eng", SolverEng) Then
              ' Lookup based on standard Solver options
7             Select Case SolverEng
              Case 1:  ' GRG Nonlinear
8                 If SolverIsAvailable(CreateSolver("Bonmin")) Then
9                     GetChosenSolver = "Bonmin"
10                End If
11            Case 2:  ' Simplex LP
12                GetChosenSolver = "CBC"
13            Case 3:  ' Evolutionary
14                If SolverIsAvailable(CreateSolver("NOMAD")) Then
15                    GetChosenSolver = "NOMAD"
16                End If
17            End Select
18        End If
          ' Make a default choice if we still don't have anything
19        If Len(GetChosenSolver) = 0 Then GetChosenSolver = GetAvailableSolvers()(LBound(GetAvailableSolvers))
20        SetChosenSolver GetChosenSolver, sheet
End Function

'/**
' * Sets the solver for an OpenSolver model.
' * @param {} SolverShortName The short name of the solver to be set
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetChosenSolver(SolverShortName As String, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
          ' Check that a valid solver has been specified
2         On Error GoTo SolverNotAllowed
3         WorksheetFunction.Match SolverShortName, GetAvailableSolvers, 0
              
4         SetNameOnSheet "OpenSolver_ChosenSolver", "=" & SolverShortName, sheet
5         Exit Sub
          
SolverNotAllowed:
6         RaiseUserError "The specified solver (" & SolverShortName & ") is not in the list of available solvers. " & _
                         "Please see the OpenSolverAPI module for the list of available solvers."
End Sub

'/**
' * Returns the objective cell in an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' * @param {} Validate If True, throws an error if the model is invalid. Defaults to True
' * @param {} RefersTo Returns the RefersTo string describing the objective
' */
Public Function GetObjectiveFunctionCell(Optional sheet As Worksheet, Optional Validate As Boolean = True, Optional RefersTo As String) As Range
1         GetActiveSheetIfMissing sheet
          
          ' Get and check the objective function
          Dim isRangeObj As Boolean, valObj As Double, ObjRefersToError As Boolean, ObjRefersToFormula As Boolean, objIsMissing As Boolean
2         GetSheetNameAsValueOrRange sheet, "solver_opt", objIsMissing, isRangeObj, GetObjectiveFunctionCell, ObjRefersToFormula, ObjRefersToError, RefersTo, valObj

3         If Validate Then
              ' If objMissing is false, but the ObjRange is empty, the objective might be an out of date reference
4             If objIsMissing = False And GetObjectiveFunctionCell Is Nothing Then
5                 RaiseUserError "OpenSolver cannot find the objective ('solver_opt' is out of date). Please re-enter the objective, and try again."
6             End If
              ' Objective is corrupted somehow
7             If ObjRefersToError Then
8                 RaiseUserError "The objective is marked #REF!, indicating this cell has been deleted. Please fix the objective, and try again."
9             End If
10        End If
End Function

'/**
' * Sets the objective cell in an OpenSolver model.
' * @param {} ObjectiveFunctionCell The cell to set
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetObjectiveFunctionCell(ObjectiveFunctionCell As Range, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         ValidateObjectiveFunctionCell ObjectiveFunctionCell
3         SetNamedRangeIfExists "solver_opt", ObjectiveFunctionCell, sheet
End Sub

'/**
' * Returns the objective sense type for an OpenSolver model. Defaults to Minimize if an invalid value is saved.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetObjectiveSense(Optional sheet As Worksheet) As ObjectiveSenseType
1         GetActiveSheetIfMissing sheet
2         GetObjectiveSense = GetNamedIntegerWithDefault(sheet, "solver_typ", ObjectiveSenseType.MinimiseObjective)
          
          ' Check that our integer is a valid value for the enum
          Dim i As Integer
3         For i = ObjectiveSenseType.[_First] To ObjectiveSenseType.[_Last]
4             If GetObjectiveSense = i Then Exit Function
5         Next i
          ' It wasn't in the enum - set default
6         GetObjectiveSense = ObjectiveSenseType.MinimiseObjective
7         SetObjectiveSense GetObjectiveSense, sheet
End Function

'/**
' * Sets the objective sense for an OpenSolver model.
' * @param {} ObjectiveSense The objective sense to set
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetObjectiveSense(ObjectiveSense As ObjectiveSenseType, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         SetIntegerNameOnSheet "solver_typ", ObjectiveSense, sheet
End Sub

'/**
' * Returns the target objective value in an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetObjectiveTargetValue(Optional sheet As Worksheet) As Double
1         GetActiveSheetIfMissing sheet
2         GetObjectiveTargetValue = GetNamedDoubleWithDefault(sheet, "solver_val", 0)
End Function

'/**
' * Sets the target objective value in an OpenSolver model.
' * @param {} ObjectiveTargetValue The target value to set
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetObjectiveTargetValue(ObjectiveTargetValue As Double, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         SetDoubleNameOnSheet "solver_val", ObjectiveTargetValue, sheet
End Sub

'/**
' * Gets the adjustable cells for an OpenSolver model, throwing an error if unset/invalid.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' * @param {} Validate If True, throws an error if the decision variables specified are missing or invalid. Defaults to True
' * @param {} RefersTo Returns the RefersTo string describing the decision variables
' */
Public Function GetDecisionVariables(Optional sheet As Worksheet, Optional Validate As Boolean = True, Optional RefersTo As String) As Range
      ' We check to see if a model exists by getting the adjustable cells. We check for a name first, as this may contain =Sheet1!$C$2:$E$2,Sheet1!#REF!
1         GetActiveSheetIfMissing sheet
          
          Dim IsRange As Boolean, value As Double, RefersToError As Boolean, RefersToFormula As Boolean, IsMissing As Boolean
2         GetSheetNameAsValueOrRange sheet, "solver_adj", IsMissing, IsRange, GetDecisionVariables, RefersToFormula, RefersToError, RefersTo, value

3         If Validate Then
4             If IsMissing Then
5                 RaiseUserError "No Solver model with decision variables was found on sheet " & sheet.Name
6             End If
              
7             If Not IsRange Then
                  ' We may not have been able to get the range because the RefersTo text of the saved range might have
                  ' exceeded Excel's range size limit. We can also try to get the named range off the sheet directly
                  ' as a fallback
8                 On Error Resume Next
9                 Set GetDecisionVariables = sheet.Range("solver_adj")
10                If Err.Number <> 0 Then
11                    RaiseUserError "A model was found on the sheet " & sheet.Name & " but the decision variable cells (" & RefersTo & ") could not be interpreted. Please redefine the decision variable cells, and try again."
12                End If
13                On Error GoTo 0
14            End If
15        End If
End Function

'/**
' * Gets the adjustable cells range (returning Nothing if invalid) for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' * @param {} DecisionVariablesRefersTo Returns the RefersTo string describing the decision variables
' */
Public Function GetDecisionVariablesWithDefault(Optional sheet As Worksheet, Optional DecisionVariablesRefersTo As String) As Range
1         GetActiveSheetIfMissing sheet
          
    On Error GoTo SetDefault:
2         Set GetDecisionVariablesWithDefault = GetDecisionVariables(sheet, True, DecisionVariablesRefersTo)
3         Exit Function
          
SetDefault:
4         Set GetDecisionVariablesWithDefault = Nothing
End Function

'/**
' * Gets the adjustable cells range (with overlap removed) for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' * @param {} DecisionVariablesRefersTo Returns the RefersTo string describing the decision variables
' */
Public Function GetDecisionVariablesNoOverlap(Optional sheet As Worksheet, Optional DecisionVariablesRefersTo As String) As Range
1         GetActiveSheetIfMissing sheet
2         Set GetDecisionVariablesNoOverlap = RemoveRangeOverlap(GetDecisionVariables(sheet, True, DecisionVariablesRefersTo))
End Function

'/**
' * Sets the adjustable cells range for an OpenSolver model.
' * @param {} DecisionVariables The range to set
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetDecisionVariables(DecisionVariables As Range, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         ValidateDecisionVariables DecisionVariables
3         SetNamedRangeIfExists "adj", DecisionVariables, sheet, True
End Sub

'/**
' * Adds a constraint in an OpenSolver model.
' * @param {} LHSRange The range to set as the constraint LHS
' * @param {} Relation The relation to set for the constraint. If Int/Bin, neither RHSRange nor RHSFormula should be set.
' * @param {} RHSRange Set if the constraint RHS is a cell/range
' * @param {} RHSFormula Set if the constraint RHS is a string formula
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub AddConstraint(LHSRange As Range, Relation As RelationConsts, Optional RHSRange As Range, Optional RHSFormula As String, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
          
          Dim NewIndex As Long
2         NewIndex = GetNumConstraints(sheet) + 1
3         UpdateConstraint NewIndex, LHSRange, Relation, RHSRange, RHSFormula, sheet
End Sub

'/**
' * Updates an existing constraint in an OpenSolver model.
' * @param {} Index The index of the constraint to update
' * @param {} LHSRange The new range to set as the constraint LHS
' * @param {} Relation The new relation to set for the constraint. If Int/Bin, neither RHSRange nor RHSFormula should be set.
' * @param {} RHSRange Set if the new constraint RHS is a cell/range
' * @param {} RHSFormula Set if the new constraint RHS is a string formula
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub UpdateConstraint(Index As Long, LHSRange As Range, Relation As RelationConsts, Optional RHSRange As Range, Optional RHSFormula As String, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
          
2         ValidateConstraint LHSRange, Relation, RHSRange, RHSFormula, sheet
          
3         SetConstraintLhs Index, LHSRange, sheet
4         SetConstraintRel Index, Relation, sheet
          
5         Select Case Relation
          Case RelationINT
6             RHSFormula = "integer"
7         Case RelationBIN
8             RHSFormula = "binary"
9         Case RelationAllDiff
10            RHSFormula = "alldiff"
11        End Select
12        If Left(RHSFormula, 1) <> "=" Then RHSFormula = "=" & RHSFormula
          
13        SetConstraintRhs Index, RHSRange, RHSFormula, sheet
          
14        If Index > GetNumConstraints(sheet) Then SetNumConstraints Index, sheet
End Sub

'/**
' * Deletes a constraint in an OpenSolver model.
' * @param {} Index The index of the constraint to delete
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub DeleteConstraint(Index As Long, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
          
          Dim NumConstraints As Long
2         NumConstraints = GetNumConstraints(sheet)
          
3         If Index > NumConstraints Or Index < 1 Then Exit Sub
          
          ' Shift all the constraints down one position
          Dim i As Long
4         For i = Index To NumConstraints - 1
              Dim LHSRange As Range, Relation As RelationConsts, RHSFormula As String, RHSRange As Range, RHSValue As Double, RHSRefersToFormula As Boolean
5             Set LHSRange = GetConstraintLhs(i + 1, sheet)
6             Relation = GetConstraintRel(i + 1, sheet)
7             Set RHSRange = GetConstraintRhs(i + 1, RHSFormula, RHSValue, RHSRefersToFormula, sheet)
8             UpdateConstraint i, LHSRange, Relation, RHSRange, RHSFormula, sheet
9         Next i
          
10        DeleteNameOnSheet "lhs" & NumConstraints, sheet, True
11        DeleteNameOnSheet "rel" & NumConstraints, sheet, True
12        DeleteNameOnSheet "rhs" & NumConstraints, sheet, True
          
13        SetNumConstraints NumConstraints - 1, sheet
End Sub

'/**
' * Clears an entire OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub ResetModel(Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
          
          Dim SolverNames() As Variant, OpenSolverNames() As Variant, Name As Variant
2         SolverNames = Array("opt", "typ", "adj", "neg", "sho", "rlx", "tol", "tim", "pre", "itr", "num", "val")
3         OpenSolverNames = Array("ChosenSolver", "DualsNewSheet", "UpdateSensitivity", "LinearityCheck", "Duals")
          
4         For Each Name In SolverNames
5             DeleteNameOnSheet CStr(Name), sheet, True
6         Next Name
7         For Each Name In OpenSolverNames
8             DeleteNameOnSheet "OpenSolver_" & CStr(Name), sheet
9         Next Name
End Sub

'/**
' * Returns the number of constraints in an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetNumConstraints(Optional sheet As Worksheet) As Long
1         GetActiveSheetIfMissing sheet
2         GetNumConstraints = GetNamedIntegerWithDefault(sheet, "solver_num", 0)
End Function

'/**
' * Sets the number of constraints in an OpenSolver model. Using Set methods to modify constraints is dangerous, it is best to use Add/Delete/UpdateConstraint.
' * @param {} NumConstraints The number of constraints to set
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetNumConstraints(NumConstraints As Long, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         SetIntegerNameOnSheet "solver_num", NumConstraints, sheet
End Sub

'/**
' * Returns the LHS range for a specified constraint in an OpenSolver model.
' * @param {} Index The index of the constraint
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' * @param {} Validate Whether to validate the LHS range. Defaults to True
' * @param {} RefersTo Returns RefersTo string representation of the LHS range
' */
Public Function GetConstraintLhs(Index As Long, Optional sheet As Worksheet, Optional Validate As Boolean = True, Optional RefersTo As String) As Range
1         GetActiveSheetIfMissing sheet
          
          Dim IsRange As Boolean, value As Double, RefersToError As Boolean, RefersToFormula As Boolean, IsMissing As Boolean
2         GetSheetNameAsValueOrRange sheet, "solver_lhs" & Index, IsMissing, IsRange, GetConstraintLhs, RefersToFormula, RefersToError, RefersTo, value
          
3         If Validate Then
              ' Must have a left hand side defined
4             If IsMissing Then
5                 RaiseUserError "The left hand side for a constraint does not appear to be defined ('solver_lhs" & Index & " is missing). Please fix this, and try again."
6             End If
              ' Must be valid
7             If RefersToError Then
8                 RaiseUserError "The constraints reference cells marked #REF!, indicating these cells have been deleted. Please fix these constraints, and try again."
9             End If
              ' LHSs must be ranges
10            If Not IsRange Then
11                RaiseUserError "A constraint was entered with a left hand side (" & RefersTo & ") that is not a range. Please update the constraint, and try again."
12            End If
13        End If
End Function

'/**
' * Sets the constraint LHS for a specified constraint in an OpenSolver model. Using Set methods to modify constraints is dangerous, it is best to use Add/Delete/UpdateConstraint.
' * @param {} Index The index of the constraint to modify
' * @param {} ConstraintLhs The cell range to set as the constraint LHS
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetConstraintLhs(Index As Long, ConstraintLhs As Range, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         SetNamedRangeIfExists "solver_lhs" & Index, ConstraintLhs, sheet
End Sub

'/**
' * Returns the relation for a specified constraint in an OpenSolver model.
' * @param {} Index The index of the constraint
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetConstraintRel(Index As Long, Optional sheet As Worksheet) As RelationConsts
1         GetActiveSheetIfMissing sheet
          
2         GetConstraintRel = GetNamedIntegerWithDefault(sheet, "solver_rel" & Index, RelationConsts.RelationLE)
          
          ' Check that our integer is a valid value for the enum
          Dim i As Integer
3         For i = RelationConsts.[_First] To RelationConsts.[_Last]
4             If GetConstraintRel = i Then Exit Function
5         Next i
          ' It wasn't in the enum - set default
6         GetConstraintRel = RelationConsts.RelationLE
7         SetConstraintRel Index, GetConstraintRel, sheet
End Function

'/**
' * Sets the constraint relation for a specified constraint in an OpenSolver model. Using Set methods to modify constraints is dangerous, it is best to use Add/Delete/UpdateConstraint.
' * @param {} Index The index of the constraint to modify
' * @param {} ConstraintRel The constraint relation to set
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetConstraintRel(Index As Long, ConstraintRel As RelationConsts, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         SetIntegerNameOnSheet "solver_rel" & Index, ConstraintRel, sheet
End Sub

'/**
' * Returns the RHS for a specified constraint in an OpenSolver model. The Formula or value parameters will be set if the RHS is not a range (in this case the function returns Nothing).
' * @param {} Index The index of the constraint
' * @param {} Formula Returns the value of the RHS if it is a string formula
' * @param {} value Returns the value of the RHS if it is a constant value
' * @param {} RefersToFormula Set to true if the RHS is a string formula
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' * @param {} Validate Whether to validate the RHS range. Defaults to True
' */
Public Function GetConstraintRhs(Index As Long, Formula As String, value As Double, RefersToFormula As Boolean, Optional sheet As Worksheet, Optional Validate As Boolean = True) As Range
1         GetActiveSheetIfMissing sheet
          
          Dim IsRange As Boolean, RefersToError As Boolean, IsMissing As Boolean
2         GetSheetNameAsValueOrRange sheet, "solver_rhs" & Index, IsMissing, IsRange, GetConstraintRhs, RefersToFormula, RefersToError, Formula, value
          
3         If Validate Then
              ' Must have a right hand side defined
4             If IsMissing Then
5                 RaiseUserError "The right hand side for a constraint does not appear to be defined ('solver_rhs" & Index & " is missing). Please fix this, and try again."
6             End If
              ' Must be valid
7             If RefersToError Then
8                 RaiseUserError "The constraints reference cells marked #REF!, indicating these cells have been deleted. Please fix these constraints, and try again."
9             End If
10        End If
End Function

'/**
' * Sets the constraint RHS for a specified constraint in an OpenSolver model. Only one of ConstraintRhsRange and ConstraintRhsFormula should be set, depending on whether the RHS is a range or a string formula. Using Set methods to modify constraints is dangerous, it is best to use Add/Delete/UpdateConstraint.
' * @param {} Index The index of the constraint to modify
' * @param {} ConstraintRhsRange Set if the constraint RHS is a cell range
' * @param {} ConstraintRhsFormula Set if the constraint RHS is a string formula
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetConstraintRhs(Index As Long, ConstraintRhsRange As Range, ConstraintRhsFormula As String, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
          
2         If ConstraintRhsRange Is Nothing Then
3             SetNameOnSheet "rhs" & Index, ConstraintRhsFormula, sheet, True
4         Else
5             SetNamedRangeIfExists "rhs" & Index, ConstraintRhsRange, sheet, True
6         End If
End Sub

'/**
' * Returns whether unconstrained variables are non-negative for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetNonNegativity(Optional sheet As Worksheet) As Boolean
1         GetActiveSheetIfMissing sheet
2         GetNonNegativity = GetNamedBooleanWithDefault(sheet, "solver_neg", True)
End Function

'/**
' * Sets whether unconstrained variables are non-negative for an OpenSolver model.
' * @param {} NonNegativity True if unconstrained variables should be non-negative
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetNonNegativity(NonNegativity As Boolean, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         SetBooleanNameOnSheet "solver_neg", NonNegativity, sheet
End Sub

'/**
' * Returns whether a post-solve linearity check will be run for an OpenSolver model
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetLinearityCheck(Optional sheet As Worksheet) As Boolean
1         GetActiveSheetIfMissing sheet
2         GetLinearityCheck = GetNamedBooleanWithDefault(sheet, "OpenSolver_LinearityCheck", True)
End Function

'/**
' * Sets the whether to run a post-solve linearity check for an OpenSolver model.
' * @param {} LinearityCheck True to run linearity check
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetLinearityCheck(LinearityCheck As Boolean, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         SetBooleanNameOnSheet "OpenSolver_LinearityCheck", LinearityCheck, sheet
End Sub

'/**
' * Returns whether to show solve progress for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetShowSolverProgress(Optional sheet As Worksheet) As Boolean
1         GetActiveSheetIfMissing sheet
2         GetShowSolverProgress = GetNamedBooleanWithDefault(sheet, "solver_sho", False)
End Function

'/**
' * Sets whether to show solve progress for an OpenSolver model.
' * @param {} ShowSolverProgress True to show progress while solving
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetShowSolverProgress(ShowSolverProgress As Boolean, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         SetBooleanNameOnSheet "solver_sho", ShowSolverProgress, sheet
End Sub

'/**
' * Returns the max solve time for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetMaxTime(Optional sheet As Worksheet) As Double
1         GetActiveSheetIfMissing sheet
2         GetMaxTime = GetNamedDoubleWithDefault(sheet, "solver_tim", MAX_LONG)
End Function

'/**
' * Sets the max solve time for an OpenSolver model.
' * @param {} MaxTime The max solve time in seconds (defaults to no limit)
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetMaxTime(Optional MaxTime As Double = MAX_LONG, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         ValidateMaxTime MaxTime
3         SetDoubleNameOnSheet "solver_tim", MaxTime, sheet
End Sub

'/**
' * Returns solver tolerance (as a double) for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetTolerance(Optional sheet As Worksheet) As Double
1         GetActiveSheetIfMissing sheet
2         GetTolerance = GetNamedDoubleWithDefault(sheet, "solver_tol", 0.05)
End Function

'/**
' * Returns solver tolerance (as a percentage) for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetToleranceAsPercentage(Optional sheet As Worksheet) As Double
1         GetActiveSheetIfMissing sheet
2         GetToleranceAsPercentage = GetTolerance(sheet) * 100
End Function

'/**
' * Sets solver tolerance for an OpenSolver model.
' * @param {} Tolerance The tolerance to set (between 0 and 1)
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetTolerance(Tolerance As Double, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         ValidateTolerance Tolerance
3         SetDoubleNameOnSheet "solver_tol", Tolerance, sheet
End Sub

'/**
' * Sets the solver tolerance (as a percentage) for an OpenSolver model.
' * @param {} Tolerance The tolerance to set as a percentage (between 0 and 100)
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetToleranceAsPercentage(Tolerance As Double, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         ValidateToleranceAsPercentage Tolerance
3         SetTolerance Tolerance / 100, sheet
End Sub

'/**
' * Returns the solver iteration limit for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetMaxIterations(Optional sheet As Worksheet) As Double
1         GetActiveSheetIfMissing sheet
2         GetMaxIterations = GetNamedDoubleWithDefault(sheet, "solver_itr", MAX_LONG)
End Function

'/**
' * Sets the solver iteration limit for an OpenSolver model.
' * @param {} MaxIterations The iteration limit to set (defaults to no limit)
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetMaxIterations(Optional MaxIterations As Double = MAX_LONG, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         ValidateMaxIterations MaxIterations
3         SetDoubleNameOnSheet "solver_itr", MaxIterations, sheet
End Sub

'/**
' * Returns the solver precision for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetPrecision(Optional sheet As Worksheet) As Double
1         GetActiveSheetIfMissing sheet
2         GetPrecision = GetNamedDoubleWithDefault(sheet, "solver_pre", 0.000001)
End Function

'/**
' * Sets the solver precision for an OpenSolver model.
' * @param {} Precision The solver precision to set
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetPrecision(Precision As Double, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         ValidatePrecision Precision
3         SetDoubleNameOnSheet "solver_pre", Precision, sheet
End Sub

'/**
' * Returns 'Extra Solver Parameters' range for specified solver in an OpenSolver model.
' * @param {} SolverShortName The short name of the solver for which parameters are being returned
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' * @param {} Validate Whether to validate the parameters range. Defaults to True
' * @param {} RefersTo Returns RefersTo string representation of the parameters range
' */
Public Function GetSolverParameters(SolverShortName As String, Optional sheet As Worksheet, Optional Validate As Boolean = True, Optional RefersTo As String) As Range
1         GetActiveSheetIfMissing sheet
          
          Dim IsRange As Boolean, value As Double, RefersToError As Boolean, RefersToFormula As Boolean, IsMissing As Boolean
2         GetSheetNameAsValueOrRange sheet, "OpenSolver_" & SolverShortName & "Parameters", IsMissing, IsRange, GetSolverParameters, RefersToFormula, RefersToError, RefersTo, value

3         If Validate Then
4             ValidateSolverParameters GetSolverParameters
5         End If
End Function

'/**
' * Sets 'Extra Parameters' range for a specified solver in an OpenSolver model.
' * @param {} SolverShortName The short name of the solver for which parameters are being set
' * @param {} SolverParameters The range containing the parameters (must be a range with two columns: keys and parameters)
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetSolverParameters(SolverShortName As String, SolverParameters As Range, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         ValidateSolverParameters SolverParameters
3         SetNamedRangeIfExists "OpenSolver_" & SolverShortName & "Parameters", SolverParameters, sheet
End Sub

'/**
' * Deletes 'Extra Parameters' range for a specified solver in an OpenSolver model.
' * @param {} SolverShortName The short name of the solver for which parameters are deleted
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub DeleteSolverParameters(SolverShortName As String, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         SetSolverParameters SolverShortName, Nothing, sheet
End Sub

'/**
' * Returns whether Solver's 'ignore integer constraints' option is set for an OpenSolver model. OpenSolver cannot solve while this option is enabled.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetIgnoreIntegerConstraints(Optional sheet As Worksheet) As Boolean
1         GetActiveSheetIfMissing sheet
2         GetIgnoreIntegerConstraints = GetNamedBooleanWithDefault(sheet, "solver_rlx", False)
End Function

'/**
' * Sets Solver's 'ignore integer constraints' option for an OpenSolver model. OpenSolver cannot solve while this option is enabled.
' * @param {} IgnoreIntegerConstraints True to turn on 'ignore integer constraints'
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetIgnoreIntegerConstraints(IgnoreIntegerConstraints As Boolean, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         SetBooleanNameOnSheet "solver_rlx", IgnoreIntegerConstraints, sheet
End Sub

'/**
' * Returns target range for sensitivity analysis output for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' * @param {} Validate If True, checks the Duals range for validity. Defaults to True
' * @param {} RefersTo Returns the RefersTo string describing the Duals range
' */
Public Function GetDuals(Optional sheet As Worksheet, Optional Validate As Boolean = True, Optional RefersTo As String) As Range
1         GetActiveSheetIfMissing sheet
          
          Dim IsRange As Boolean, value As Double, RefersToError As Boolean, RefersToFormula As Boolean, IsMissing As Boolean
2         GetSheetNameAsValueOrRange sheet, "OpenSolver_Duals", IsMissing, IsRange, GetDuals, RefersToFormula, RefersToError, RefersTo, value

3         If Validate Then
              ' TODO
4         End If
End Function

'/**
' * Sets target range for sensitivity analysis output for an OpenSolver model.
' * @param {} Duals The target range for output (Nothing for no sensitivity analysis)
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetDuals(Duals As Range, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         ValidateDuals Duals
3         SetNamedRangeIfExists "OpenSolver_Duals", Duals, sheet
End Sub

'/**
' * Returns whether 'Output sensitivity analysis' is set for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetDualsOnSheet(Optional sheet As Worksheet) As Boolean
1         GetActiveSheetIfMissing sheet
2         GetDualsOnSheet = GetNamedBooleanWithDefault(sheet, "OpenSolver_DualsNewSheet", False)
End Function

'/**
' * Sets the value of 'Output sensitivity analysis' for an OpenSolver model.
' * @param {} DualsOnSheet True to set 'Output sensitivity analysis'
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetDualsOnSheet(DualsOnSheet As Boolean, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         SetBooleanNameOnSheet "OpenSolver_DualsNewSheet", DualsOnSheet, sheet
End Sub

'/**
' * Returns True if 'Output sensitivity analysis' destination is set to 'updating any previous sheet' for an OpenSolver model, and False if set to 'on a new sheet'.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetUpdateSensitivity(Optional sheet As Worksheet) As Boolean
1         GetActiveSheetIfMissing sheet
2         GetUpdateSensitivity = GetNamedBooleanWithDefault(sheet, "OpenSolver_UpdateSensitivity", True)
End Function

'/**
' * Sets the destination option for 'Output sensitivity analysis' for an OpenSolver model.
' * @param {} UpdateSensitivity True to set 'updating any previous sheet'. False to set 'on a new sheet'
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetUpdateSensitivity(UpdateSensitivity As Boolean, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         SetBooleanNameOnSheet "OpenSolver_UpdateSensitivity", UpdateSensitivity, sheet
End Sub

'/**
' * Gets the QuickSolve parameter range for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' * @param {} Validate If True, an error will be thrown if no range is set
' * @param {} RefersTo Returns RefersTo string representation of the parameters range
' */
Public Function GetQuickSolveParameters(Optional sheet As Worksheet, Optional Validate As Boolean = True, Optional RefersTo As String) As Range
1         GetActiveSheetIfMissing sheet
          
          Dim IsRange As Boolean, value As Double, RefersToError As Boolean, RefersToFormula As Boolean, IsMissing As Boolean
2         GetSheetNameAsValueOrRange sheet, "OpenSolverModelParameters", IsMissing, IsRange, GetQuickSolveParameters, RefersToFormula, RefersToError, RefersTo, value
          
3         If Validate Then
4             If GetQuickSolveParameters Is Nothing Then
5                 RaiseUserError "No parameter range could be found on the worksheet. Please use ""Initialize Quick Solve Parameters""" & _
                                 "to define the cells that you wish to change between successive OpenSolver solves. Note that changes " & _
                                 "to these cells must lead to changes in the underlying model's right hand side values for its constraints."
6             End If
7         End If
End Function

'/**
' * Sets the QuickSolve parameter range for an OpenSolver model.
' * @param {} QuickSolveParameters The parameter range to set
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetQuickSolveParameters(QuickSolveParameters As Range, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         ValidateQuickSolveParameters QuickSolveParameters, sheet
3         SetNamedRangeIfExists "OpenSolverModelParameters", QuickSolveParameters, sheet
End Sub

'/**
' * Initializes QuickSolve procedure for an OpenSolver model.
' * @param {} SolveRelaxation If True, all integer and boolean constraints will be relaxed to allow continuous values for these variables. Defaults to False
' * @param {} MinimiseUserInteraction If True, all dialogs and messages will be suppressed. Use this when automating a lot of solves so that there are no interruptions. Defaults to False
' * @param {} LinearityCheckOffset Sets the base value used for checking if the model is linear. Change this if a non-linear model is not being detected as non-linear. Defaults to 0
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub InitializeQuickSolve(Optional SolveRelaxation As Boolean = False, Optional MinimiseUserInteraction As Boolean = False, Optional LinearityCheckOffset As Double = 0, Optional sheet As Worksheet)
1         ClearError
2         On Error GoTo ErrorHandler
          
3         GetActiveSheetIfMissing sheet

4         If Not CreateSolver(GetChosenSolver(sheet)).ModelType = Diff Then
5             RaiseUserError "The selected solver does not support QuickSolve"
6         End If

          Dim ParamRange As Range
7         Set ParamRange = GetQuickSolveParameters(sheet, Validate:=True)  ' Throws error if missing
8         Set QuickSolver = New COpenSolver
9         QuickSolver.BuildModelFromSolverData LinearityCheckOffset, GetLinearityCheck(sheet), MinimiseUserInteraction, SolveRelaxation, sheet
10        QuickSolver.InitializeQuickSolve ParamRange

ExitSub:
11        Exit Sub

ErrorHandler:
12        ReportError "OpenSolverAPI", "InitializeQuickSolve", True, MinimiseUserInteraction
13        GoTo ExitSub
End Sub

'/**
' * Runs a QuickSolve for currently initialized QuickSolve model.
' * @param {} MinimiseUserInteraction If True, all dialogs and messages will be suppressed. Use this when automating a lot of solves so that there are no interruptions. Defaults to False
' */
Public Function RunQuickSolve(Optional SolveRelaxation As Boolean = False, Optional MinimiseUserInteraction As Boolean = False) As OpenSolverResult
1         ClearError
2         On Error GoTo ErrorHandler
          
          Dim InteractiveStatus As Boolean
3         InteractiveStatus = Application.Interactive
4         Application.Interactive = False

5         If QuickSolver Is Nothing Then
6             RaiseUserError "There is no model to solve, and so the quick solve cannot be completed. Please select the Initialize Quick Solve command."
7         Else
8             QuickSolver.DoQuickSolve SolveRelaxation, MinimiseUserInteraction
9             RunQuickSolve = QuickSolver.SolveStatus
10        End If

11        If Not MinimiseUserInteraction Then QuickSolver.ReportAnySolutionSubOptimality

ExitFunction:
12        Application.Interactive = InteractiveStatus
13        Exit Function

ErrorHandler:
14        ReportError "OpenSolverMain", "RunQuickSolve", True, MinimiseUserInteraction
15        If OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError Then
16            RunQuickSolve = AbortedThruUserAction
17        Else
18            RunQuickSolve = OpenSolverResult.ErrorOccurred
19        End If
20        GoTo ExitFunction
End Function

'/**
' * Clears any initialized QuickSolve.
' */
Public Sub ClearQuickSolve()
1         Set QuickSolver = Nothing
End Sub

'/**
' * Returns the RefersTo string for the objective in an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetObjectiveFunctionCellRefersTo(Optional sheet As Worksheet) As String
1         GetActiveSheetIfMissing sheet
2         GetObjectiveFunctionCell sheet:=sheet, Validate:=False, RefersTo:=GetObjectiveFunctionCellRefersTo
End Function

'/**
' * Sets the objective cell using a RefersTo string in an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} ObjectiveRefersTo The RefersTo string to set as the objective
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetObjectiveFunctionCellRefersTo(ObjectiveFunctionCellRefersTo As String, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         ValidateObjectiveFunctionCellRefersTo ObjectiveFunctionCellRefersTo
3         SetRefersToNameOnSheet "solver_opt", ObjectiveFunctionCellRefersTo, sheet
End Sub

'/**
' * Returns the RefersTo string for the decision variables in an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetDecisionVariablesRefersTo(Optional sheet As Worksheet) As String
1         GetActiveSheetIfMissing sheet
2         GetDecisionVariables sheet:=sheet, Validate:=False, RefersTo:=GetDecisionVariablesRefersTo
End Function

'/**
' * Sets the adjustable cells using a RefersTo string for an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} DecisionVariablesRefersTo The RefersTo string describing the decision variable range to set
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetDecisionVariablesRefersTo(DecisionVariablesRefersTo As String, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         ValidateDecisionVariablesRefersTo DecisionVariablesRefersTo
3         SetRefersToNameOnSheet "adj", DecisionVariablesRefersTo, sheet, True
End Sub

'/**
' * Updates an existing constraint in an OpenSolver model using RefersTo strings. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} Index The index of the constraint to update
' * @param {} LHSRefersTo The new RefersTo string to set as the constraint LHS
' * @param {} Relation The new relation to set for the constraint. If Int/Bin, neither RHSRange nor RHSFormula should be set.
' * @param {} RHSRefersTo The new RefersTo string to set as the constraint RHS
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub UpdateConstraintRefersTo(Index As Long, LHSRefersTo As String, Relation As RelationConsts, RHSRefersTo As String, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
          
2         ValidateConstraintRefersTo LHSRefersTo, Relation, RHSRefersTo, sheet
          
3         SetConstraintLhsRefersTo Index, LHSRefersTo, sheet
4         SetConstraintRel Index, Relation, sheet
          
5         Select Case Relation
          Case RelationINT
6             RHSRefersTo = "integer"
7         Case RelationBIN
8             RHSRefersTo = "binary"
9         Case RelationAllDiff
10            RHSRefersTo = "alldiff"
11        End Select
          
12        SetConstraintRhsRefersTo Index, RHSRefersTo, sheet
          
13        If Index > GetNumConstraints(sheet) Then SetNumConstraints Index, sheet
End Sub

'/**
' * Gets the constraint description in RefersTo format for the specified constraint in an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} Index The index of the constraint
' * @param {} LHSRefersTo Returns the RefersTo string describing the constraint LHS
' * @param {} Relation Returns the constraint relation type
' * @param {} RHSRefersTo Returns the RefersTo string describing the constraint RHS
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub GetConstraintRefersTo(Index As Long, LHSRefersTo As String, Relation As RelationConsts, RHSRefersTo As String, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
          
2         LHSRefersTo = GetConstraintLhsRefersTo(Index, sheet)
3         Relation = GetConstraintRel(Index, sheet)
4         If RelationHasRHS(Relation) Then
5             RHSRefersTo = GetConstraintRhsRefersTo(Index, sheet)
6         Else
7             RHSRefersTo = vbNullString
8         End If
End Sub

'/**
' * Returns the RefersTo string for the LHS of the specified constraint in an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} Index The index of the constraint
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetConstraintLhsRefersTo(Index As Long, Optional sheet As Worksheet) As String
1         GetActiveSheetIfMissing sheet
2         GetConstraintLhs Index, sheet:=sheet, Validate:=False, RefersTo:=GetConstraintLhsRefersTo
End Function

'/**
' * Sets the constraint LHS using a RefersTo string for a specified constraint in an OpenSolver model. Using Set methods to modify constraints is dangerous, it is best to use Add/Delete/UpdateConstraint. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} Index The index of the constraint to modify
' * @param {} ConstraintLhsRefersTo The RefersTo string to set as the constraint LHS
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetConstraintLhsRefersTo(Index As Long, ConstraintLhsRefersTo As String, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         SetRefersToNameOnSheet "solver_lhs" & Index, ConstraintLhsRefersTo, sheet
End Sub

'/**
' * Returns the RefersTo string for the LHS of the specified constraint in an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} Index The index of the constraint
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetConstraintRhsRefersTo(Index As Long, Optional sheet As Worksheet) As String
1         GetActiveSheetIfMissing sheet
          
          Dim value As Double, RefersToFormula As Boolean
2         GetConstraintRhs Index, GetConstraintRhsRefersTo, value, RefersToFormula, sheet:=sheet, Validate:=False
End Function

'/**
' * Sets the constraint RHS using a RefersTo string for a specified constraint in an OpenSolver model. Using Set methods to modify constraints is dangerous, it is best to use Add/Delete/UpdateConstraint. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} Index The index of the constraint to modify
' * @param {} ConstraintRhsRefersTo The RefersTo string to set as the constraint RHS
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetConstraintRhsRefersTo(Index As Long, ConstraintRhsRefersTo As String, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         SetRefersToNameOnSheet "solver_rhs" & Index, ConstraintRhsRefersTo, sheet
End Sub

'/**
' * Returns the RefersTo string for the 'Extra Solver Parameters' range for specified solver in an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} SolverShortName The short name of the solver for which parameters are being returned
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetSolverParametersRefersTo(SolverShortName As String, Optional sheet As Worksheet) As String
1         GetActiveSheetIfMissing sheet
2         GetSolverParameters SolverShortName, sheet:=sheet, Validate:=False, RefersTo:=GetSolverParametersRefersTo
End Function

'/**
' * Sets the 'Extra Parameters' range using a RefersTo string for a specified solver in an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} Index The index of the constraint to modify
' * @param {} SolverParametersRefersTo The RefersTo string to set as the extra parameters range
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetSolverParametersRefersTo(SolverShortName As String, SolverParametersRefersTo As String, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
          
2         ValidateSolverParametersRefersTo SolverParametersRefersTo
3         SetRefersToNameOnSheet "OpenSolver_" & SolverShortName & "Parameters", SolverParametersRefersTo, sheet
End Sub

'/**
' * Returns the RefersTo string for the sensitivity analysis output in an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetDualsRefersTo(Optional sheet As Worksheet) As String
1         GetActiveSheetIfMissing sheet
2         GetDuals sheet:=sheet, Validate:=False, RefersTo:=GetDualsRefersTo
End Function

'/**
' * Sets target range for sensitivity analysis output using a RefersTo string for an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} DualsRefersTo The RefersTo string describing the target range for output (Nothing for no sensitivity analysis)
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetDualsRefersTo(DualsRefersTo As String, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         ValidateDualsRefersTo DualsRefersTo
3         SetRefersToNameOnSheet "OpenSolver_Duals", DualsRefersTo, sheet
End Sub

'/**
' * Returns the RefersTo string for the QuickSolve parameter range in an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetQuickSolveParametersRefersTo(Optional sheet As Worksheet) As String
1         GetActiveSheetIfMissing sheet
2         GetQuickSolveParameters sheet:=sheet, Validate:=False, RefersTo:=GetQuickSolveParametersRefersTo
End Function

'/**
' * Sets the QuickSolve parameter range using a RefersTo string for an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} QuickSolveParametersRefersTo The RefersTo string describing the parameter range to set
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetQuickSolveParametersRefersTo(QuickSolveParametersRefersTo As String, Optional sheet As Worksheet)
1         GetActiveSheetIfMissing sheet
2         ValidateQuickSolveParametersRefersTo QuickSolveParametersRefersTo, sheet
3         SetRefersToNameOnSheet "OpenSolverModelParameters", QuickSolveParametersRefersTo, sheet
End Sub

