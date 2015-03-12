Attribute VB_Name = "OpenSolverAPI"
Option Explicit

Public Function GetAvailableSolvers() As Variant()
    GetAvailableSolvers = Array("CBC", "Gurobi", "NeosCBC", "Bonmin", "Couenne", "NOMAD", "NeosBon", "NeosCou")
End Function

Public Function GetChosenSolver(Optional book As Workbook, Optional sheet As Worksheet) As String
    GetActiveBookAndSheetIfMissing book, sheet
    If Not GetNameValueIfExists(book, EscapeSheetName(sheet) & "OpenSolver_ChosenSolver", GetChosenSolver) Then
        GoTo SetDefault
    End If
    
    ' Check solver is an allowed solver
    On Error GoTo SetDefault
    WorksheetFunction.Match GetChosenSolver, GetAvailableSolvers, 0
    Exit Function
    
SetDefault:
    GetChosenSolver = GetAvailableSolvers()(0)
    SetChosenSolver GetChosenSolver, book, sheet
End Function

Public Sub SetChosenSolver(Solver As String, Optional book As Workbook, Optional sheet As Worksheet)
    ' Check that a valid solver has been specified
    On Error GoTo SolverNotAllowed
    WorksheetFunction.Match Solver, GetAvailableSolvers, 0
        
    SetNameOnSheet "OpenSolver_ChosenSolver", "=" & Solver, book, sheet
    Exit Sub
    
SolverNotAllowed:
    Err.Raise OpenSolver_ModelError, Description:="The specified solver (" & Solver & ") is not in the list of available solvers. " & _
                                                  "Please see the OpenSolverAPI module for the list of available solvers."
End Sub

Public Function GetDualsNewSheet(Optional book As Workbook, Optional sheet As Worksheet) As Boolean
    GetDualsNewSheet = GetNamedBooleanWithDefault("OpenSolver_DualsNewSheet", book, sheet, False)
End Function

Public Sub SetDualsNewSheet(DualsNewSheet As Boolean, Optional book As Workbook, Optional sheet As Worksheet)
    SetBooleanNameOnSheet "OpenSolver_DualsNewSheet", DualsNewSheet, book, sheet
End Sub

Public Function GetUpdateSensitivity(Optional book As Workbook, Optional sheet As Worksheet) As Boolean
    GetUpdateSensitivity = GetNamedBooleanWithDefault("OpenSolver_UpdateSensitivity", book, sheet, True)
End Function

Public Sub SetUpdateSensitivity(UpdateSensitivity As Boolean, Optional book As Workbook, Optional sheet As Worksheet)
    SetBooleanNameOnSheet "OpenSolver_UpdateSensitivity", UpdateSensitivity, book, sheet
End Sub

Public Function GetLinearityCheck(Optional book As Workbook, Optional sheet As Worksheet) As Boolean
    GetLinearityCheck = GetNamedIntegerAsBooleanWithDefault("OpenSolver_LinearityCheck", book, sheet, True)
End Function

Public Sub SetLinearityCheck(LinearityCheck As Boolean, Optional book As Workbook, Optional sheet As Worksheet)
    SetBooleanAsIntegerNameOnSheet "OpenSolver_LinearityCheck", LinearityCheck, book, sheet
End Sub

Public Function GetDuals(Optional book As Workbook, Optional sheet As Worksheet) As Range
    If Not GetNamedRangeIfExistsOnSheet(sheet, "OpenSolver_Duals", GetDuals) Then Set GetDuals = Nothing
End Function

Public Sub SetDuals(Duals As Range, Optional book As Workbook, Optional sheet As Worksheet)
    SetNamedRangeIfExists "OpenSolver_Duals", Duals, book, sheet
End Sub

Public Function GetSolverParameters(Solver As String, Optional book As Workbook, Optional sheet As Worksheet) As Range
    If Not GetNamedRangeIfExistsOnSheet(sheet, "OpenSolver_" & Solver & "Parameters", GetSolverParameters) Then Set GetSolverParameters = Nothing
End Function

Public Sub SetSolverParameters(Solver As String, SolverParameters As Range, Optional book As Workbook, Optional sheet As Worksheet)
    SetNamedRangeIfExists "OpenSolver_" & Solver & "Parameters", SolverParameters, book, sheet
End Sub

Public Function GetQuickSolveParameters(Optional book As Workbook, Optional sheet As Worksheet) As Range
    If Not GetNamedRangeIfExistsOnSheet(sheet, "OpenSolverModelParameters", GetQuickSolveParameters) Then Set GetQuickSolveParameters = Nothing
End Function

Public Sub SetQuickSolveParameters(QuickSolveParameters As Range, Optional book As Workbook, Optional sheet As Worksheet)
    SetNamedRangeIfExists "OpenSolverModelParameters", QuickSolveParameters, book, sheet
End Sub

Public Function GetDecisionVariables(Optional book As Workbook, Optional sheet As Worksheet) As Range
    GetActiveBookAndSheetIfMissing book, sheet
              
    ' We check to see if a model exists by getting the adjustable cells. We check for a name first, as this may contain =Sheet1!$C$2:$E$2,Sheet1!#REF!
    Dim n As Name
    If Not NameExistsInWorkbook(book, EscapeSheetName(sheet) & "solver_adj", n) Then
        Err.Raise Number:=OpenSolver_ModelError, Description:="No Solver model with decision variables was found on sheet " & sheet.Name
    End If
    
    GetNamedRangeIfExistsOnSheet sheet, "solver_adj", GetDecisionVariables
    If GetDecisionVariables Is Nothing Then
        Err.Raise OpenSolver_ModelError, Description:="A model was found on the sheet " & sheet.Name & " but the decision variable cells (" & n & ") could not be interpreted. Please redefine the decision variable cells, and try again."
    End If
End Function

Public Function GetDecisionVariablesWithDefault(Optional book As Workbook, Optional sheet As Worksheet) As Range
    On Error GoTo SetDefault:
    Set GetDecisionVariablesWithDefault = GetDecisionVariables(book, sheet)
    Exit Function
    
SetDefault:
    Set GetDecisionVariablesWithDefault = Nothing
End Function

Public Function GetDecisionVariablesNoOverlap(Optional book As Workbook, Optional sheet As Worksheet) As Range
    Set GetDecisionVariablesNoOverlap = RemoveRangeOverlap(GetDecisionVariables(book, sheet))
End Function

Public Sub SetDecisionVariables(DecisionVariables As Range, Optional book As Workbook, Optional sheet As Worksheet)
    SetSolverNamedRangeIfExists "adj", DecisionVariables, book, sheet
End Sub

Public Function GetNonNegativity(Optional book As Workbook, Optional sheet As Worksheet) As Boolean
    GetNonNegativity = GetNamedIntegerAsBooleanWithDefault("solver_neg", book, sheet, True)
End Function

Public Sub SetNonNegativity(NonNegativity As Boolean, Optional book As Workbook, Optional sheet As Worksheet)
    SetBooleanAsIntegerNameOnSheet "solver_neg", NonNegativity, book, sheet
End Sub

Public Function GetShowSolverProgress(Optional book As Workbook, Optional sheet As Worksheet) As Boolean
    GetShowSolverProgress = GetNamedIntegerAsBooleanWithDefault("solver_sho", book, sheet, False)
End Function

Public Sub SetShowSolverProgress(ShowSolverProgress As Boolean, Optional book As Workbook, Optional sheet As Worksheet)
    SetBooleanAsIntegerNameOnSheet "solver_sho", ShowSolverProgress, book, sheet
End Sub

Public Function GetIgnoreIntegerConstraints(Optional book As Workbook, Optional sheet As Worksheet) As Boolean
    GetIgnoreIntegerConstraints = GetNamedIntegerAsBooleanWithDefault("solver_rlx", book, sheet, False)
End Function

Public Sub SetIgnoreIntegerConstraints(IgnoreIntegerConstraints As Boolean, Optional book As Workbook, Optional sheet As Worksheet)
    SetBooleanAsIntegerNameOnSheet "solver_rlx", IgnoreIntegerConstraints, book, sheet
End Sub

Public Function GetTolerance(Optional book As Workbook, Optional sheet As Worksheet) As Double
    GetTolerance = GetNamedDoubleWithDefault("solver_tol", book, sheet, 0.05)
End Function

Public Function GetToleranceAsPercentage(Optional book As Workbook, Optional sheet As Worksheet) As Double
    GetToleranceAsPercentage = GetTolerance(book, sheet) * 100
End Function

Public Sub SetTolerance(Tolerance As Double, Optional book As Workbook, Optional sheet As Worksheet)
    SetDoubleNameOnSheet "solver_tol", Tolerance, book, sheet
End Sub

Public Sub SetToleranceAsPercentage(Tolerance As Double, Optional book As Workbook, Optional sheet As Worksheet)
    SetTolerance Tolerance / 100, book, sheet
End Sub

Public Function GetMaxTime(Optional book As Workbook, Optional sheet As Worksheet) As Long
    GetMaxTime = GetNamedIntegerWithDefault("solver_tim", book, sheet, 999999999)
End Function

Public Sub SetMaxTime(MaxTime As Long, Optional book As Workbook, Optional sheet As Worksheet)
    SetIntegerNameOnSheet "solver_tim", MaxTime, book, sheet
End Sub

Public Function GetPrecision(Optional book As Workbook, Optional sheet As Worksheet) As Double
    GetPrecision = GetNamedDoubleWithDefault("solver_pre", book, sheet, 0.000001)
End Function

Public Sub SetPrecision(Precision As Double, Optional book As Workbook, Optional sheet As Worksheet)
    SetDoubleNameOnSheet "solver_pre", Precision, book, sheet
End Sub

Public Function GetMaxIterations(Optional book As Workbook, Optional sheet As Worksheet) As Long
    GetMaxIterations = GetNamedIntegerWithDefault("solver_itr", book, sheet, 100)
End Function

Public Sub SetMaxIterations(MaxIterations As Long, Optional book As Workbook, Optional sheet As Worksheet)
    SetIntegerNameOnSheet "solver_itr", MaxIterations, book, sheet
End Sub

Public Function GetObjectiveSense(Optional book As Workbook, Optional sheet As Worksheet) As ObjectiveSenseType
    GetObjectiveSense = GetNamedIntegerWithDefault("solver_typ", book, sheet, ObjectiveSenseType.MinimiseObjective)
    
    ' Check that our integer is a valid value for the enum
    Dim i As Integer
    For i = ObjectiveSenseType.[_First] To ObjectiveSenseType.[_Last]
        If GetObjectiveSense = i Then Exit Function
    Next i
    ' It wasn't in the enum - set default
    GetObjectiveSense = ObjectiveSenseType.MinimiseObjective
    SetObjectiveSense GetObjectiveSense, book, sheet
End Function

Public Sub SetObjectiveSense(ObjectiveSense As ObjectiveSenseType, Optional book As Workbook, Optional sheet As Worksheet)
    SetIntegerNameOnSheet "solver_typ", ObjectiveSense, book, sheet
End Sub

Public Function GetNumConstraints(Optional book As Workbook, Optional sheet As Worksheet) As Long
    GetNumConstraints = GetNamedIntegerWithDefault("solver_num", book, sheet, 0)
End Function

Public Sub SetNumConstraints(NumConstraints As Long, Optional book As Workbook, Optional sheet As Worksheet)
    SetIntegerNameOnSheet "solver_num", NumConstraints, book, sheet
End Sub

Public Function GetObjectiveTargetValue(Optional book As Workbook, Optional sheet As Worksheet) As Double
    GetObjectiveTargetValue = GetNamedDoubleWithDefault("solver_val", book, sheet, 0)
End Function

Public Sub SetObjectiveTargetValue(ObjectiveTargetValue As Double, Optional book As Workbook, Optional sheet As Worksheet)
    SetDoubleNameOnSheet "solver_val", ObjectiveTargetValue, book, sheet
End Sub

Public Function GetObjectiveFunctionCell(Optional book As Workbook, Optional sheet As Worksheet, Optional ValidateObjective As Boolean = False) As Range
    GetActiveBookAndSheetIfMissing book, sheet
    
    ' Get and check the objective function
    Dim isRangeObj As Boolean, valObj As Double, ObjRefersToError As Boolean, ObjRefersToFormula As Boolean, sRefersToObj As String, objIsMissing As Boolean
    GetNameAsValueOrRange book, EscapeSheetName(sheet) & "solver_opt", objIsMissing, isRangeObj, GetObjectiveFunctionCell, ObjRefersToFormula, ObjRefersToError, sRefersToObj, valObj

    If Not ValidateObjective Or GetObjectiveFunctionCell Is Nothing Then Exit Function

    ' If objMissing is false, but the ObjRange is empty, the objective might be an out of date reference
    If objIsMissing = False And GetObjectiveFunctionCell Is Nothing Then
        Err.Raise Number:=OpenSolver_BuildError, Description:="OpenSolver cannot find the objective ('solver_opt' is out of date). Please re-enter the objective, and try again."
    End If
    ' Objective is corrupted somehow
    If ObjRefersToError Then
        Err.Raise Number:=OpenSolver_BuildError, Description:="The objective is marked #REF!, indicating this cell has been deleted. Please fix the objective, and try again."
    End If
    ' Objective has a value that is not a number
    If VarType(GetObjectiveFunctionCell.Value2) <> vbDouble Then
        If VarType(GetObjectiveFunctionCell.Value2) = vbError Then
            Err.Raise Number:=OpenSolver_BuildError, Description:="The objective cell appears to contain an error. This could have occurred if there is a divide by zero error or if you have used the wrong function (eg #DIV/0! or #VALUE!). Please fix this, and try again."
        Else
            Err.Raise Number:=OpenSolver_BuildError, Description:="The objective cell does not appear to contain a numeric value. Please fix this, and try again."
        End If
    End If
End Function

Public Function GetObjectiveFunctionCellWithValidation(Optional book As Workbook, Optional sheet As Worksheet) As Range
    Set GetObjectiveFunctionCellWithValidation = GetObjectiveFunctionCell(book, sheet, True)
End Function

Public Sub SetObjectiveFunctionCell(ObjectiveFunctionCell As Range, Optional book As Workbook, Optional sheet As Worksheet)
    SetNamedRangeIfExists "solver_opt", ObjectiveFunctionCell, book, sheet
End Sub

Public Function GetConstraintRel(Index As Long, Optional book As Workbook, Optional sheet As Worksheet) As RelationConsts
    GetConstraintRel = GetNamedIntegerWithDefault("solver_rel" & Index, book, sheet, RelationConsts.RelationLE)
    
    ' Check that our integer is a valid value for the enum
    Dim i As Integer
    For i = RelationConsts.[_First] To RelationConsts.[_Last]
        If GetConstraintRel = i Then Exit Function
    Next i
    ' It wasn't in the enum - set default
    GetConstraintRel = RelationConsts.RelationLE
    SetConstraintRel Index, GetConstraintRel, book, sheet
End Function

Public Sub SetConstraintRel(Index As Long, ConstraintRel As RelationConsts, Optional book As Workbook, Optional sheet As Worksheet)
    SetIntegerNameOnSheet "solver_rel" & Index, ConstraintRel, book, sheet
End Sub

Public Function GetConstraintLhs(Index As Long, Optional book As Workbook, Optional sheet As Worksheet) As Range
    GetActiveBookAndSheetIfMissing book, sheet
    
    Set GetConstraintLhs = Nothing
    
    Dim IsRange As Boolean, value As Double, RefersToError As Boolean, RefersToFormula As Boolean, RangeFormula As String, IsMissing As Boolean
    GetNameAsValueOrRange book, EscapeSheetName(sheet) & "solver_lhs" & Index, IsMissing, IsRange, GetConstraintLhs, RefersToFormula, RefersToError, RangeFormula, value
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
        Err.Raise Number:=OpenSolver_BuildError, Description:="A constraint was entered with a left hand side (" & RangeFormula & ") that is not a range. Please update the constraint, and try again."
    End If
End Function

Public Sub SetConstraintLhs(Index As Long, ConstraintLhs As Range, Optional book As Workbook, Optional sheet As Worksheet)
    SetNamedRangeIfExists "solver_lhs" & Index, ConstraintLhs, book, sheet
End Sub

Public Function GetConstraintRhs(Index As Long, Formula As String, value As Double, RefersToFormula As Boolean, Optional book As Workbook, Optional sheet As Worksheet) As Range
    GetActiveBookAndSheetIfMissing book, sheet
    
    Set GetConstraintRhs = Nothing
    
    Dim IsRange As Boolean, RefersToError As Boolean, IsMissing As Boolean
    GetNameAsValueOrRange book, EscapeSheetName(sheet) & "solver_rhs" & Index, IsMissing, IsRange, GetConstraintRhs, RefersToFormula, RefersToError, Formula, value
    ' Must have a right hand side defined
    If IsMissing Then
        Err.Raise Number:=OpenSolver_BuildError, Description:="The right hand side for a constraint does not appear to be defined ('solver_rhs" & Index & " is missing). Please fix this, and try again."
    End If
    ' Must be valid
    If RefersToError Then
        Err.Raise Number:=OpenSolver_BuildError, Description:="The constraints reference cells marked #REF!, indicating these cells have been deleted. Please fix these constraints, and try again."
    End If
End Function

Public Sub SetConstraintRhs(Index As Long, ConstraintRhsRange As Range, ConstraintRhsFormula As String, Optional book As Workbook, Optional sheet As Worksheet)
    If ConstraintRhsRange Is Nothing Then
        SetSolverNameOnSheet "rhs", ConstraintRhsFormula, book, sheet
    Else
        SetNamedRangeIfExists "solver_rhs" & Index, ConstraintRhsRange, book, sheet
    End If
End Sub
