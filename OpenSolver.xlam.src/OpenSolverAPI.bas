Attribute VB_Name = "OpenSolverAPI"
Option Explicit

Public Const sOpenSolverVersion As String = "2.8.2"
Public Const sOpenSolverDate As String = "2016.02.15"

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
    CheckLocationValid  ' Check for unicode in path
    
    ClearError
    On Error GoTo ErrorHandler
    
    Dim InteractiveStatus As Boolean
    InteractiveStatus = Application.Interactive
    Application.Interactive = False
    
    GetActiveSheetIfMissing sheet

    RunOpenSolver = OpenSolverResult.Unsolved
    Dim OpenSolver As COpenSolver
    Set OpenSolver = New COpenSolver

    OpenSolver.BuildModelFromSolverData LinearityOffset, GetLinearityCheck(sheet), MinimiseUserInteraction, SolveRelaxation, sheet
    ' Only proceed with solve if nothing detected while building model
    If OpenSolver.SolveStatus = OpenSolverResult.Unsolved Then
        SolveModel OpenSolver, SolveRelaxation, MinimiseUserInteraction
    End If
    
    RunOpenSolver = OpenSolver.SolveStatus
    If Not MinimiseUserInteraction Then OpenSolver.ReportAnySolutionSubOptimality

ExitFunction:
    Application.Interactive = InteractiveStatus
    Set OpenSolver = Nothing    ' Free any OpenSolver memory used
    Exit Function

ErrorHandler:
    ReportError "OpenSolverAPI", "RunOpenSolver", True, MinimiseUserInteraction
    If OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError Then
        RunOpenSolver = AbortedThruUserAction
    Else
        RunOpenSolver = OpenSolverResult.ErrorOccurred
    End If
    GoTo ExitFunction
End Function

'/**
' * Gets a list of short names for all solvers that can be set
' */
Public Function GetAvailableSolvers() As String()
    GetAvailableSolvers = StringArray("CBC", "Gurobi", "NeosCBC", "Bonmin", "Couenne", "NOMAD", "NeosBon", "NeosCou")
End Function

'/**
' * Gets the short name of the currently selected solver for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetChosenSolver(Optional sheet As Worksheet) As String
    GetActiveSheetIfMissing sheet
    If Not GetNamedStringIfExists(sheet, "OpenSolver_ChosenSolver", GetChosenSolver) Then GoTo SetDefault
    
    ' Check solver is an allowed solver
    On Error GoTo SetDefault
    WorksheetFunction.Match GetChosenSolver, GetAvailableSolvers, 0
    Exit Function
    
SetDefault:
    GetChosenSolver = GetAvailableSolvers()(LBound(GetAvailableSolvers))
    SetChosenSolver GetChosenSolver, sheet
End Function

'/**
' * Sets the solver for an OpenSolver model.
' * @param {} SolverShortName The short name of the solver to be set
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetChosenSolver(SolverShortName As String, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    ' Check that a valid solver has been specified
    On Error GoTo SolverNotAllowed
    WorksheetFunction.Match SolverShortName, GetAvailableSolvers, 0
        
    SetNameOnSheet "OpenSolver_ChosenSolver", "=" & SolverShortName, sheet
    Exit Sub
    
SolverNotAllowed:
    Err.Raise OpenSolver_ModelError, Description:="The specified solver (" & SolverShortName & ") is not in the list of available solvers. " & _
                                                  "Please see the OpenSolverAPI module for the list of available solvers."
End Sub

'/**
' * Returns the objective cell in an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' * @param {} Validate If True, throws an error if the model is invalid. Defaults to True
' * @param {} RefersTo Returns the RefersTo string describing the objective
' */
Public Function GetObjectiveFunctionCell(Optional sheet As Worksheet, Optional Validate As Boolean = True, Optional RefersTo As String) As Range
    GetActiveSheetIfMissing sheet
    
    ' Get and check the objective function
    Dim isRangeObj As Boolean, valObj As Double, ObjRefersToError As Boolean, ObjRefersToFormula As Boolean, objIsMissing As Boolean
    GetSheetNameAsValueOrRange sheet, "solver_opt", objIsMissing, isRangeObj, GetObjectiveFunctionCell, ObjRefersToFormula, ObjRefersToError, RefersTo, valObj

    If Validate Then
        ' If objMissing is false, but the ObjRange is empty, the objective might be an out of date reference
        If objIsMissing = False And GetObjectiveFunctionCell Is Nothing Then
            Err.Raise Number:=OpenSolver_BuildError, Description:="OpenSolver cannot find the objective ('solver_opt' is out of date). Please re-enter the objective, and try again."
        End If
        ' Objective is corrupted somehow
        If ObjRefersToError Then
            Err.Raise Number:=OpenSolver_BuildError, Description:="The objective is marked #REF!, indicating this cell has been deleted. Please fix the objective, and try again."
        End If
        
        If GetObjectiveFunctionCell Is Nothing Then Exit Function
        
        ' Objective has a value that is not a number
        If VarType(GetObjectiveFunctionCell.Value2) <> vbDouble Then
            If VarType(GetObjectiveFunctionCell.Value2) = vbError Then
                Err.Raise Number:=OpenSolver_BuildError, Description:="The objective cell appears to contain an error. This could have occurred if there is a divide by zero error or if you have used the wrong function (eg #DIV/0! or #VALUE!). Please fix this, and try again."
            Else
                Err.Raise Number:=OpenSolver_BuildError, Description:="The objective cell does not appear to contain a numeric value. Please fix this, and try again."
            End If
        End If
    End If
End Function

'/**
' * Sets the objective cell in an OpenSolver model.
' * @param {} ObjectiveFunctionCell The cell to set
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetObjectiveFunctionCell(ObjectiveFunctionCell As Range, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    ValidateObjectiveFunctionCell ObjectiveFunctionCell
    SetNamedRangeIfExists "solver_opt", ObjectiveFunctionCell, sheet
End Sub

'/**
' * Returns the objective sense type for an OpenSolver model. Defaults to Minimize if an invalid value is saved.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetObjectiveSense(Optional sheet As Worksheet) As ObjectiveSenseType
    GetActiveSheetIfMissing sheet
    GetObjectiveSense = GetNamedIntegerWithDefault(sheet, "solver_typ", ObjectiveSenseType.MinimiseObjective)
    
    ' Check that our integer is a valid value for the enum
    Dim i As Integer
    For i = ObjectiveSenseType.[_First] To ObjectiveSenseType.[_Last]
        If GetObjectiveSense = i Then Exit Function
    Next i
    ' It wasn't in the enum - set default
    GetObjectiveSense = ObjectiveSenseType.MinimiseObjective
    SetObjectiveSense GetObjectiveSense, sheet
End Function

'/**
' * Sets the objective sense for an OpenSolver model.
' * @param {} ObjectiveSense The objective sense to set
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetObjectiveSense(ObjectiveSense As ObjectiveSenseType, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    SetIntegerNameOnSheet "solver_typ", ObjectiveSense, sheet
End Sub

'/**
' * Returns the target objective value in an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetObjectiveTargetValue(Optional sheet As Worksheet) As Double
    GetActiveSheetIfMissing sheet
    GetObjectiveTargetValue = GetNamedDoubleWithDefault(sheet, "solver_val", 0)
End Function

'/**
' * Sets the target objective value in an OpenSolver model.
' * @param {} ObjectiveTargetValue The target value to set
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetObjectiveTargetValue(ObjectiveTargetValue As Double, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    SetDoubleNameOnSheet "solver_val", ObjectiveTargetValue, sheet
End Sub

'/**
' * Gets the adjustable cells for an OpenSolver model, throwing an error if unset/invalid.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' * @param {} Validate If True, throws an error if the decision variables specified are missing or invalid. Defaults to True
' * @param {} RefersTo Returns the RefersTo string describing the decision variables
' */
Public Function GetDecisionVariables(Optional sheet As Worksheet, Optional Validate As Boolean = True, Optional RefersTo As String) As Range
' We check to see if a model exists by getting the adjustable cells. We check for a name first, as this may contain =Sheet1!$C$2:$E$2,Sheet1!#REF!
    GetActiveSheetIfMissing sheet
    
    Dim IsRange As Boolean, value As Double, RefersToError As Boolean, RefersToFormula As Boolean, IsMissing As Boolean
    GetSheetNameAsValueOrRange sheet, "solver_adj", IsMissing, IsRange, GetDecisionVariables, RefersToFormula, RefersToError, RefersTo, value

    If Validate Then
        If IsMissing Then
            Err.Raise Number:=OpenSolver_ModelError, Description:="No Solver model with decision variables was found on sheet " & sheet.Name
        End If
        
        If Not IsRange Then
            Err.Raise OpenSolver_ModelError, Description:="A model was found on the sheet " & sheet.Name & " but the decision variable cells (" & RefersTo & ") could not be interpreted. Please redefine the decision variable cells, and try again."
        End If
    End If
End Function

'/**
' * Gets the adjustable cells range (returning Nothing if invalid) for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' * @param {} DecisionVariablesRefersTo Returns the RefersTo string describing the decision variables
' */
Public Function GetDecisionVariablesWithDefault(Optional sheet As Worksheet, Optional DecisionVariablesRefersTo As String) As Range
    GetActiveSheetIfMissing sheet
    
    On Error GoTo SetDefault:
    Set GetDecisionVariablesWithDefault = GetDecisionVariables(sheet, True, DecisionVariablesRefersTo)
    Exit Function
    
SetDefault:
    Set GetDecisionVariablesWithDefault = Nothing
End Function

'/**
' * Gets the adjustable cells range (with overlap removed) for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' * @param {} DecisionVariablesRefersTo Returns the RefersTo string describing the decision variables
' */
Public Function GetDecisionVariablesNoOverlap(Optional sheet As Worksheet, Optional DecisionVariablesRefersTo As String) As Range
    GetActiveSheetIfMissing sheet
    Set GetDecisionVariablesNoOverlap = RemoveRangeOverlap(GetDecisionVariables(sheet, True, DecisionVariablesRefersTo))
End Function

'/**
' * Sets the adjustable cells range for an OpenSolver model.
' * @param {} DecisionVariables The range to set
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetDecisionVariables(DecisionVariables As Range, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    ValidateDecisionVariables DecisionVariables
    SetNamedRangeIfExists "adj", DecisionVariables, sheet, True
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
    GetActiveSheetIfMissing sheet
    
    Dim NewIndex As Long
    NewIndex = GetNumConstraints(sheet) + 1
    UpdateConstraint NewIndex, LHSRange, Relation, RHSRange, RHSFormula, sheet
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
    GetActiveSheetIfMissing sheet
    
    ValidateConstraint LHSRange, Relation, RHSRange, RHSFormula, sheet
    
    SetConstraintLhs Index, LHSRange, sheet
    SetConstraintRel Index, Relation, sheet
    
    Select Case Relation
    Case RelationINT
        RHSFormula = "integer"
    Case RelationBIN
        RHSFormula = "binary"
    Case RelationAllDiff
        RHSFormula = "alldiff"
    End Select
    If Left(RHSFormula, 1) <> "=" Then RHSFormula = "=" & RHSFormula
    
    SetConstraintRhs Index, RHSRange, RHSFormula, sheet
    
    If Index > GetNumConstraints(sheet) Then SetNumConstraints Index, sheet
End Sub

'/**
' * Deletes a constraint in an OpenSolver model.
' * @param {} Index The index of the constraint to delete
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub DeleteConstraint(Index As Long, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    
    Dim NumConstraints As Long
    NumConstraints = GetNumConstraints(sheet)
    
    If Index > NumConstraints Or Index < 1 Then Exit Sub
    
    ' Shift all the constraints down one position
    Dim i As Long
    For i = Index To NumConstraints - 1
        Dim LHSRange As Range, Relation As RelationConsts, RHSFormula As String, RHSRange As Range, RHSValue As Double, RHSRefersToFormula As Boolean
        Set LHSRange = GetConstraintLhs(i + 1, sheet)
        Relation = GetConstraintRel(i + 1, sheet)
        Set RHSRange = GetConstraintRhs(i + 1, RHSFormula, RHSValue, RHSRefersToFormula, sheet)
        UpdateConstraint i, LHSRange, Relation, RHSRange, RHSFormula, sheet
    Next i
    
    DeleteNameOnSheet "lhs" & NumConstraints, sheet, True
    DeleteNameOnSheet "rel" & NumConstraints, sheet, True
    DeleteNameOnSheet "rhs" & NumConstraints, sheet, True
    
    SetNumConstraints NumConstraints - 1, sheet
End Sub

'/**
' * Clears an entire OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub ResetModel(Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    
    Dim SolverNames() As Variant, OpenSolverNames() As Variant, Name As Variant
    SolverNames = Array("opt", "typ", "adj", "neg", "sho", "rlx", "tol", "tim", "pre", "itr", "num", "val")
    OpenSolverNames = Array("ChosenSolver", "DualsNewSheet", "UpdateSensitivity", "LinearityCheck", "Duals")
    
    For Each Name In SolverNames
        DeleteNameOnSheet CStr(Name), sheet, True
    Next Name
    For Each Name In OpenSolverNames
        DeleteNameOnSheet "OpenSolver_" & CStr(Name), sheet
    Next Name
End Sub

'/**
' * Returns the number of constraints in an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetNumConstraints(Optional sheet As Worksheet) As Long
    GetActiveSheetIfMissing sheet
    GetNumConstraints = GetNamedIntegerWithDefault(sheet, "solver_num", 0)
End Function

'/**
' * Sets the number of constraints in an OpenSolver model. Using Set methods to modify constraints is dangerous, it is best to use Add/Delete/UpdateConstraint.
' * @param {} NumConstraints The number of constraints to set
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetNumConstraints(NumConstraints As Long, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    SetIntegerNameOnSheet "solver_num", NumConstraints, sheet
End Sub

'/**
' * Returns the LHS range for a specified constraint in an OpenSolver model.
' * @param {} Index The index of the constraint
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' * @param {} Validate Whether to validate the LHS range. Defaults to True
' * @param {} RefersTo Returns RefersTo string representation of the LHS range
' */
Public Function GetConstraintLhs(Index As Long, Optional sheet As Worksheet, Optional Validate As Boolean = True, Optional RefersTo As String) As Range
    GetActiveSheetIfMissing sheet
    
    Dim IsRange As Boolean, value As Double, RefersToError As Boolean, RefersToFormula As Boolean, IsMissing As Boolean
    GetSheetNameAsValueOrRange sheet, "solver_lhs" & Index, IsMissing, IsRange, GetConstraintLhs, RefersToFormula, RefersToError, RefersTo, value
    
    If Validate Then
        ' Must have a left hand side defined
        If IsMissing Then
            Err.Raise Number:=OpenSolver_BuildError, Description:="The left hand side for a constraint does not appear to be defined ('solver_lhs" & Index & " is missing). Please fix this, and try again."
        End If
        ' Must be valid
        If RefersToError Then
            Err.Raise Number:=OpenSolver_BuildError, Description:="The constraints reference cells marked #REF!, indicating these cells have been deleted. Please fix these constraints, and try again."
        End If
        ' LHSs must be ranges
        If Not IsRange Then
            Err.Raise Number:=OpenSolver_BuildError, Description:="A constraint was entered with a left hand side (" & RefersTo & ") that is not a range. Please update the constraint, and try again."
        End If
    End If
End Function

'/**
' * Sets the constraint LHS for a specified constraint in an OpenSolver model. Using Set methods to modify constraints is dangerous, it is best to use Add/Delete/UpdateConstraint.
' * @param {} Index The index of the constraint to modify
' * @param {} ConstraintLhs The cell range to set as the constraint LHS
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetConstraintLhs(Index As Long, ConstraintLhs As Range, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    SetNamedRangeIfExists "solver_lhs" & Index, ConstraintLhs, sheet
End Sub

'/**
' * Returns the relation for a specified constraint in an OpenSolver model.
' * @param {} Index The index of the constraint
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetConstraintRel(Index As Long, Optional sheet As Worksheet) As RelationConsts
    GetActiveSheetIfMissing sheet
    
    GetConstraintRel = GetNamedIntegerWithDefault(sheet, "solver_rel" & Index, RelationConsts.RelationLE)
    
    ' Check that our integer is a valid value for the enum
    Dim i As Integer
    For i = RelationConsts.[_First] To RelationConsts.[_Last]
        If GetConstraintRel = i Then Exit Function
    Next i
    ' It wasn't in the enum - set default
    GetConstraintRel = RelationConsts.RelationLE
    SetConstraintRel Index, GetConstraintRel, sheet
End Function

'/**
' * Sets the constraint relation for a specified constraint in an OpenSolver model. Using Set methods to modify constraints is dangerous, it is best to use Add/Delete/UpdateConstraint.
' * @param {} Index The index of the constraint to modify
' * @param {} ConstraintRel The constraint relation to set
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetConstraintRel(Index As Long, ConstraintRel As RelationConsts, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    SetIntegerNameOnSheet "solver_rel" & Index, ConstraintRel, sheet
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
    GetActiveSheetIfMissing sheet
    
    Dim IsRange As Boolean, RefersToError As Boolean, IsMissing As Boolean
    GetSheetNameAsValueOrRange sheet, "solver_rhs" & Index, IsMissing, IsRange, GetConstraintRhs, RefersToFormula, RefersToError, Formula, value
    
    If Validate Then
        ' Must have a right hand side defined
        If IsMissing Then
            Err.Raise Number:=OpenSolver_BuildError, Description:="The right hand side for a constraint does not appear to be defined ('solver_rhs" & Index & " is missing). Please fix this, and try again."
        End If
        ' Must be valid
        If RefersToError Then
            Err.Raise Number:=OpenSolver_BuildError, Description:="The constraints reference cells marked #REF!, indicating these cells have been deleted. Please fix these constraints, and try again."
        End If
    End If
End Function

'/**
' * Sets the constraint RHS for a specified constraint in an OpenSolver model. Only one of ConstraintRhsRange and ConstraintRhsFormula should be set, depending on whether the RHS is a range or a string formula. Using Set methods to modify constraints is dangerous, it is best to use Add/Delete/UpdateConstraint.
' * @param {} Index The index of the constraint to modify
' * @param {} ConstraintRhsRange Set if the constraint RHS is a cell range
' * @param {} ConstraintRhsFormula Set if the constraint RHS is a string formula
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetConstraintRhs(Index As Long, ConstraintRhsRange As Range, ConstraintRhsFormula As String, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    
    If ConstraintRhsRange Is Nothing Then
        SetNameOnSheet "rhs" & Index, ConstraintRhsFormula, sheet, True
    Else
        SetNamedRangeIfExists "rhs" & Index, ConstraintRhsRange, sheet, True
    End If
End Sub

'/**
' * Returns whether unconstrained variables are non-negative for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetNonNegativity(Optional sheet As Worksheet) As Boolean
    GetActiveSheetIfMissing sheet
    GetNonNegativity = GetNamedBooleanWithDefault(sheet, "solver_neg", True)
End Function

'/**
' * Sets whether unconstrained variables are non-negative for an OpenSolver model.
' * @param {} NonNegativity True if unconstrained variables should be non-negative
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetNonNegativity(NonNegativity As Boolean, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    SetBooleanNameOnSheet "solver_neg", NonNegativity, sheet
End Sub

'/**
' * Returns whether a post-solve linearity check will be run for an OpenSolver model
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetLinearityCheck(Optional sheet As Worksheet) As Boolean
    GetActiveSheetIfMissing sheet
    GetLinearityCheck = GetNamedBooleanWithDefault(sheet, "OpenSolver_LinearityCheck", True)
End Function

'/**
' * Sets the whether to run a post-solve linearity check for an OpenSolver model.
' * @param {} LinearityCheck True to run linearity check
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetLinearityCheck(LinearityCheck As Boolean, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    SetBooleanNameOnSheet "OpenSolver_LinearityCheck", LinearityCheck, sheet
End Sub

'/**
' * Returns whether to show solve progress for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetShowSolverProgress(Optional sheet As Worksheet) As Boolean
    GetActiveSheetIfMissing sheet
    GetShowSolverProgress = GetNamedBooleanWithDefault(sheet, "solver_sho", False)
End Function

'/**
' * Sets whether to show solve progress for an OpenSolver model.
' * @param {} ShowSolverProgress True to show progress while solving
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetShowSolverProgress(ShowSolverProgress As Boolean, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    SetBooleanNameOnSheet "solver_sho", ShowSolverProgress, sheet
End Sub

'/**
' * Returns the max solve time for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetMaxTime(Optional sheet As Worksheet) As Double
    GetActiveSheetIfMissing sheet
    GetMaxTime = GetNamedDoubleWithDefault(sheet, "solver_tim", 999999999)
End Function

'/**
' * Sets the max solve time for an OpenSolver model.
' * @param {} MaxTime The max solve time in seconds
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetMaxTime(MaxTime As Double, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    SetDoubleNameOnSheet "solver_tim", MaxTime, sheet
End Sub

'/**
' * Returns solver tolerance (as a double) for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetTolerance(Optional sheet As Worksheet) As Double
    GetActiveSheetIfMissing sheet
    GetTolerance = GetNamedDoubleWithDefault(sheet, "solver_tol", 0.05)
End Function

'/**
' * Returns solver tolerance (as a percentage) for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetToleranceAsPercentage(Optional sheet As Worksheet) As Double
    GetActiveSheetIfMissing sheet
    GetToleranceAsPercentage = GetTolerance(sheet) * 100
End Function

'/**
' * Sets solver tolerance for an OpenSolver model.
' * @param {} Tolerance The tolerance to set (between 0 and 1)
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetTolerance(Tolerance As Double, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    SetDoubleNameOnSheet "solver_tol", Tolerance, sheet
End Sub

'/**
' * Sets the solver tolerance (as a percentage) for an OpenSolver model.
' * @param {} Tolerance The tolerance to set as a percentage (between 0 and 100)
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetToleranceAsPercentage(Tolerance As Double, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    SetTolerance Tolerance / 100, sheet
End Sub

'/**
' * Returns the solver iteration limit for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetMaxIterations(Optional sheet As Worksheet) As Long
    GetActiveSheetIfMissing sheet
    GetMaxIterations = GetNamedIntegerWithDefault(sheet, "solver_itr", 999999999)
End Function

'/**
' * Sets the solver iteration limit for an OpenSolver model.
' * @param {} MaxIterations The iteration limit to set
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetMaxIterations(MaxIterations As Long, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    SetIntegerNameOnSheet "solver_itr", MaxIterations, sheet
End Sub

'/**
' * Returns the solver precision for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetPrecision(Optional sheet As Worksheet) As Double
    GetActiveSheetIfMissing sheet
    GetPrecision = GetNamedDoubleWithDefault(sheet, "solver_pre", 0.000001)
End Function

'/**
' * Sets the solver precision for an OpenSolver model.
' * @param {} Precision The solver precision to set
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetPrecision(Precision As Double, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    SetDoubleNameOnSheet "solver_pre", Precision, sheet
End Sub

'/**
' * Returns 'Extra Solver Parameters' range for specified solver in an OpenSolver model.
' * @param {} SolverShortName The short name of the solver for which parameters are being returned
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' * @param {} Validate Whether to validate the parameters range. Defaults to True
' * @param {} RefersTo Returns RefersTo string representation of the parameters range
' */
Public Function GetSolverParameters(SolverShortName As String, Optional sheet As Worksheet, Optional Validate As Boolean = True, Optional RefersTo As String) As Range
    GetActiveSheetIfMissing sheet
    
    Dim IsRange As Boolean, value As Double, RefersToError As Boolean, RefersToFormula As Boolean, IsMissing As Boolean
    GetSheetNameAsValueOrRange sheet, "OpenSolver_" & SolverShortName & "Parameters", IsMissing, IsRange, GetSolverParameters, RefersToFormula, RefersToError, RefersTo, value

    If Validate Then
        ValidateSolverParameters GetSolverParameters
    End If
End Function

'/**
' * Sets 'Extra Parameters' range for a specified solver in an OpenSolver model.
' * @param {} SolverShortName The short name of the solver for which parameters are being set
' * @param {} SolverParameters The range containing the parameters (must be a range with two columns: keys and parameters)
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetSolverParameters(SolverShortName As String, SolverParameters As Range, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    ValidateSolverParameters SolverParameters
    SetNamedRangeIfExists "OpenSolver_" & SolverShortName & "Parameters", SolverParameters, sheet
End Sub

'/**
' * Deletes 'Extra Parameters' range for a specified solver in an OpenSolver model.
' * @param {} SolverShortName The short name of the solver for which parameters are deleted
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub DeleteSolverParameters(SolverShortName As String, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    SetSolverParameters SolverShortName, Nothing, sheet
End Sub

'/**
' * Returns whether Solver's 'ignore integer constraints' option is set for an OpenSolver model. OpenSolver cannot solve while this option is enabled.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetIgnoreIntegerConstraints(Optional sheet As Worksheet) As Boolean
    GetActiveSheetIfMissing sheet
    GetIgnoreIntegerConstraints = GetNamedBooleanWithDefault(sheet, "solver_rlx", False)
End Function

'/**
' * Sets Solver's 'ignore integer constraints' option for an OpenSolver model. OpenSolver cannot solve while this option is enabled.
' * @param {} IgnoreIntegerConstraints True to turn on 'ignore integer constraints'
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetIgnoreIntegerConstraints(IgnoreIntegerConstraints As Boolean, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    SetBooleanNameOnSheet "solver_rlx", IgnoreIntegerConstraints, sheet
End Sub

'/**
' * Returns target range for sensitivity analysis output for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' * @param {} Validate If True, checks the Duals range for validity. Defaults to True
' * @param {} RefersTo Returns the RefersTo string describing the Duals range
' */
Public Function GetDuals(Optional sheet As Worksheet, Optional Validate As Boolean = True, Optional RefersTo As String) As Range
    GetActiveSheetIfMissing sheet
    
    Dim IsRange As Boolean, value As Double, RefersToError As Boolean, RefersToFormula As Boolean, IsMissing As Boolean
    GetSheetNameAsValueOrRange sheet, "OpenSolver_Duals", IsMissing, IsRange, GetDuals, RefersToFormula, RefersToError, RefersTo, value

    If Validate Then
        ' TODO
    End If
End Function

'/**
' * Sets target range for sensitivity analysis output for an OpenSolver model.
' * @param {} Duals The target range for output (Nothing for no sensitivity analysis)
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetDuals(Duals As Range, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    ValidateDuals Duals
    SetNamedRangeIfExists "OpenSolver_Duals", Duals, sheet
End Sub

'/**
' * Returns whether 'Output sensitivity analysis' is set for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetDualsOnSheet(Optional sheet As Worksheet) As Boolean
    GetActiveSheetIfMissing sheet
    GetDualsOnSheet = GetNamedBooleanWithDefault(sheet, "OpenSolver_DualsNewSheet", False)
End Function

'/**
' * Sets the value of 'Output sensitivity analysis' for an OpenSolver model.
' * @param {} DualsOnSheet True to set 'Output sensitivity analysis'
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetDualsOnSheet(DualsOnSheet As Boolean, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    SetBooleanNameOnSheet "OpenSolver_DualsNewSheet", DualsOnSheet, sheet
End Sub

'/**
' * Returns True if 'Output sensitivity analysis' destination is set to 'updating any previous sheet' for an OpenSolver model, and False if set to 'on a new sheet'.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetUpdateSensitivity(Optional sheet As Worksheet) As Boolean
    GetActiveSheetIfMissing sheet
    GetUpdateSensitivity = GetNamedBooleanWithDefault(sheet, "OpenSolver_UpdateSensitivity", True)
End Function

'/**
' * Sets the destination option for 'Output sensitivity analysis' for an OpenSolver model.
' * @param {} UpdateSensitivity True to set 'updating any previous sheet'. False to set 'on a new sheet'
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetUpdateSensitivity(UpdateSensitivity As Boolean, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    SetBooleanNameOnSheet "OpenSolver_UpdateSensitivity", UpdateSensitivity, sheet
End Sub

'/**
' * Gets the QuickSolve parameter range for an OpenSolver model.
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' * @param {} Validate If True, an error will be thrown if no range is set
' * @param {} RefersTo Returns RefersTo string representation of the parameters range
' */
Public Function GetQuickSolveParameters(Optional sheet As Worksheet, Optional Validate As Boolean = True, Optional RefersTo As String) As Range
    GetActiveSheetIfMissing sheet
    
    Dim IsRange As Boolean, value As Double, RefersToError As Boolean, RefersToFormula As Boolean, IsMissing As Boolean
    GetSheetNameAsValueOrRange sheet, "OpenSolverModelParameters", IsMissing, IsRange, GetQuickSolveParameters, RefersToFormula, RefersToError, RefersTo, value
    
    If Validate Then
        If GetQuickSolveParameters Is Nothing Then
            Err.Raise OpenSolver_BuildError, Description:="No parameter range could be found on the worksheet. Please use ""Initialize Quick Solve Parameters""" & _
                                                          "to define the cells that you wish to change between successive OpenSolver solves. Note that changes " & _
                                                          "to these cells must lead to changes in the underlying model's right hand side values for its constraints."
        End If
    End If
End Function

'/**
' * Sets the QuickSolve parameter range for an OpenSolver model.
' * @param {} QuickSolveParameters The parameter range to set
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetQuickSolveParameters(QuickSolveParameters As Range, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    ValidateQuickSolveParameters QuickSolveParameters, sheet
    SetNamedRangeIfExists "OpenSolverModelParameters", QuickSolveParameters, sheet
End Sub

'/**
' * Initializes QuickSolve procedure for an OpenSolver model.
' * @param {} SolveRelaxation If True, all integer and boolean constraints will be relaxed to allow continuous values for these variables. Defaults to False
' * @param {} MinimiseUserInteraction If True, all dialogs and messages will be suppressed. Use this when automating a lot of solves so that there are no interruptions. Defaults to False
' * @param {} LinearityCheckOffset Sets the base value used for checking if the model is linear. Change this if a non-linear model is not being detected as non-linear. Defaults to 0
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub InitializeQuickSolve(Optional SolveRelaxation As Boolean = False, Optional MinimiseUserInteraction As Boolean = False, Optional LinearityCheckOffset As Double = 0, Optional sheet As Worksheet)
    ClearError
    On Error GoTo ErrorHandler
    
    GetActiveSheetIfMissing sheet

    If Not CreateSolver(GetChosenSolver(sheet)).ModelType = Diff Then
        Err.Raise OpenSolver_ModelError, Description:="The selected solver does not support QuickSolve"
    End If

    Dim ParamRange As Range
    Set ParamRange = GetQuickSolveParameters(sheet, Validate:=True)  ' Throws error if missing
    Set QuickSolver = New COpenSolver
    QuickSolver.BuildModelFromSolverData LinearityCheckOffset, GetLinearityCheck(sheet), MinimiseUserInteraction, SolveRelaxation, sheet
    QuickSolver.InitializeQuickSolve ParamRange

ExitSub:
    Exit Sub

ErrorHandler:
    ReportError "OpenSolverAPI", "InitializeQuickSolve", True, MinimiseUserInteraction
    GoTo ExitSub
End Sub

'/**
' * Runs a QuickSolve for currently initialized QuickSolve model.
' * @param {} MinimiseUserInteraction If True, all dialogs and messages will be suppressed. Use this when automating a lot of solves so that there are no interruptions. Defaults to False
' */
Public Function RunQuickSolve(Optional SolveRelaxation As Boolean = False, Optional MinimiseUserInteraction As Boolean = False) As OpenSolverResult
    ClearError
    On Error GoTo ErrorHandler

    If QuickSolver Is Nothing Then
        Err.Raise OpenSolver_SolveError, Description:="There is no model to solve, and so the quick solve cannot be completed. Please select the Initialize Quick Solve command."
    Else
        QuickSolver.DoQuickSolve SolveRelaxation, MinimiseUserInteraction
        RunQuickSolve = QuickSolver.SolveStatus
    End If

    If Not MinimiseUserInteraction Then QuickSolver.ReportAnySolutionSubOptimality

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

'/**
' * Clears any initialized QuickSolve.
' */
Public Sub ClearQuickSolve()
    Set QuickSolver = Nothing
End Sub

'/**
' * Returns the RefersTo string for the objective in an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetObjectiveFunctionCellRefersTo(Optional sheet As Worksheet) As String
    GetActiveSheetIfMissing sheet
    GetObjectiveFunctionCell sheet:=sheet, Validate:=False, RefersTo:=GetObjectiveFunctionCellRefersTo
End Function

'/**
' * Sets the objective cell using a RefersTo string in an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} ObjectiveRefersTo The RefersTo string to set as the objective
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetObjectiveFunctionCellRefersTo(ObjectiveFunctionCellRefersTo As String, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    ValidateObjectiveFunctionCellRefersTo ObjectiveFunctionCellRefersTo
    SetRefersToNameOnSheet "solver_opt", ObjectiveFunctionCellRefersTo, sheet
End Sub

'/**
' * Returns the RefersTo string for the decision variables in an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetDecisionVariablesRefersTo(Optional sheet As Worksheet) As String
    GetActiveSheetIfMissing sheet
    GetDecisionVariables sheet:=sheet, Validate:=False, RefersTo:=GetDecisionVariablesRefersTo
End Function

'/**
' * Sets the adjustable cells using a RefersTo string for an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} DecisionVariablesRefersTo The RefersTo string describing the decision variable range to set
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetDecisionVariablesRefersTo(DecisionVariablesRefersTo As String, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    ValidateDecisionVariablesRefersTo DecisionVariablesRefersTo
    SetRefersToNameOnSheet "adj", DecisionVariablesRefersTo, sheet, True
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
    GetActiveSheetIfMissing sheet
    
    ValidateConstraintRefersTo LHSRefersTo, Relation, RHSRefersTo, sheet
    
    SetConstraintLhsRefersTo Index, LHSRefersTo, sheet
    SetConstraintRel Index, Relation, sheet
    
    Select Case Relation
    Case RelationINT
        RHSRefersTo = "integer"
    Case RelationBIN
        RHSRefersTo = "binary"
    Case RelationAllDiff
        RHSRefersTo = "alldiff"
    End Select
    
    SetConstraintRhsRefersTo Index, RHSRefersTo, sheet
    
    If Index > GetNumConstraints(sheet) Then SetNumConstraints Index, sheet
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
    GetActiveSheetIfMissing sheet
    
    LHSRefersTo = GetConstraintLhsRefersTo(Index, sheet)
    Relation = GetConstraintRel(Index, sheet)
    If RelationHasRHS(Relation) Then
        RHSRefersTo = GetConstraintRhsRefersTo(Index, sheet)
    Else
        RHSRefersTo = vbNullString
    End If
End Sub

'/**
' * Returns the RefersTo string for the LHS of the specified constraint in an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} Index The index of the constraint
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetConstraintLhsRefersTo(Index As Long, Optional sheet As Worksheet) As String
    GetActiveSheetIfMissing sheet
    GetConstraintLhs Index, sheet:=sheet, Validate:=False, RefersTo:=GetConstraintLhsRefersTo
End Function

'/**
' * Sets the constraint LHS using a RefersTo string for a specified constraint in an OpenSolver model. Using Set methods to modify constraints is dangerous, it is best to use Add/Delete/UpdateConstraint. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} Index The index of the constraint to modify
' * @param {} ConstraintLhsRefersTo The RefersTo string to set as the constraint LHS
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetConstraintLhsRefersTo(Index As Long, ConstraintLhsRefersTo As String, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    SetRefersToNameOnSheet "solver_lhs" & Index, ConstraintLhsRefersTo, sheet
End Sub

'/**
' * Returns the RefersTo string for the LHS of the specified constraint in an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} Index The index of the constraint
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetConstraintRhsRefersTo(Index As Long, Optional sheet As Worksheet) As String
    GetActiveSheetIfMissing sheet
    
    Dim value As Double, RefersToFormula As Boolean
    GetConstraintRhs Index, GetConstraintRhsRefersTo, value, RefersToFormula, sheet:=sheet, Validate:=False
End Function

'/**
' * Sets the constraint RHS using a RefersTo string for a specified constraint in an OpenSolver model. Using Set methods to modify constraints is dangerous, it is best to use Add/Delete/UpdateConstraint. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} Index The index of the constraint to modify
' * @param {} ConstraintRhsRefersTo The RefersTo string to set as the constraint RHS
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetConstraintRhsRefersTo(Index As Long, ConstraintRhsRefersTo As String, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    SetRefersToNameOnSheet "solver_rhs" & Index, ConstraintRhsRefersTo, sheet
End Sub

'/**
' * Returns the RefersTo string for the 'Extra Solver Parameters' range for specified solver in an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} SolverShortName The short name of the solver for which parameters are being returned
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetSolverParametersRefersTo(SolverShortName As String, Optional sheet As Worksheet) As String
    GetActiveSheetIfMissing sheet
    GetSolverParameters SolverShortName, sheet:=sheet, Validate:=False, RefersTo:=GetSolverParametersRefersTo
End Function

'/**
' * Sets the 'Extra Parameters' range using a RefersTo string for a specified solver in an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} Index The index of the constraint to modify
' * @param {} SolverParametersRefersTo The RefersTo string to set as the extra parameters range
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetSolverParametersRefersTo(SolverShortName As String, SolverParametersRefersTo As String, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    
    ValidateSolverParametersRefersTo SolverParametersRefersTo
    SetRefersToNameOnSheet "OpenSolver_" & SolverShortName & "Parameters", SolverParametersRefersTo, sheet
End Sub

'/**
' * Returns the RefersTo string for the sensitivity analysis output in an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetDualsRefersTo(Optional sheet As Worksheet) As String
    GetActiveSheetIfMissing sheet
    GetDuals sheet:=sheet, Validate:=False, RefersTo:=GetDualsRefersTo
End Function

'/**
' * Sets target range for sensitivity analysis output using a RefersTo string for an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} DualsRefersTo The RefersTo string describing the target range for output (Nothing for no sensitivity analysis)
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetDualsRefersTo(DualsRefersTo As String, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    ValidateDualsRefersTo DualsRefersTo
    SetRefersToNameOnSheet "OpenSolver_Duals", DualsRefersTo, sheet
End Sub

'/**
' * Returns the RefersTo string for the QuickSolve parameter range in an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Function GetQuickSolveParametersRefersTo(Optional sheet As Worksheet) As String
    GetActiveSheetIfMissing sheet
    GetQuickSolveParameters sheet:=sheet, Validate:=False, RefersTo:=GetQuickSolveParametersRefersTo
End Function

'/**
' * Sets the QuickSolve parameter range using a RefersTo string for an OpenSolver model. WARNING: Do not use RefersTo methods unless you know what you are doing!
' * @param {} QuickSolveParametersRefersTo The RefersTo string describing the parameter range to set
' * @param {} sheet The worksheet containing the model (defaults to active worksheet)
' */
Public Sub SetQuickSolveParametersRefersTo(QuickSolveParametersRefersTo As String, Optional sheet As Worksheet)
    GetActiveSheetIfMissing sheet
    ValidateQuickSolveParametersRefersTo QuickSolveParametersRefersTo, sheet
    SetRefersToNameOnSheet "OpenSolverModelParameters", QuickSolveParametersRefersTo, sheet
End Sub

