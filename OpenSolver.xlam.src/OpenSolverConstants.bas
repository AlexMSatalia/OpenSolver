Attribute VB_Name = "OpenSolverConstants"
Option Explicit

Public Const OpenSolverRegName = "OpenSolver"
Public Const PreferencesRegName = "Preferences"
' From rondebruin.nl:
' The GetSetting default argument can't be an emptystring on Mac
Public Const VALUE_IF_MISSING As String = "?"

Public Const EPSILON As Double = 0.00001
Public Const MAX_LONG As Long = 2147483647
Public Const MAX_DOUBLE As Double = 1E+307 ' Actually ~1.8E+308 but errors occur if we try to set cell values that high

'Solution results, as reported by Excel Solver
' FROM http://msdn.microsoft.com/en-us/library/ff197237.aspx
' 0 Solver found a solution. All constraints and optimality conditions are satisfied.
' 1 Solver has converged to the current solution. All constraints are satisfied.
' 2 Solver cannot improve the current solution. All constraints are satisfied.
' 3 Stop chosen when the maximum iteration limit was reached.
' 4 The Objective Cell values do not converge.
' 5 Solver could not find a feasible solution.
' 6 Solver stopped at user's request.
' 7 The linearity conditions required by this LP Solver are not satisfied.
' 8 The problem is too large for Solver to handle.
' 9 Solver encountered an error value in a target or constraint cell.
' 10 Stop chosen when the maximum time limit was reached.
' 11 There is not enough memory available to solve the problem.
' 13 Error in model. Please verify that all cells and constraints are valid.
' 14 Solver found an integer solution within tolerance. All constraints are satisfied.
' 15 Stop chosen when the maximum number of feasible [integer] solutions was reached.
' 16 Stop chosen when the maximum number of feasible [integer] subproblems was reached.
' 17 Solver converged in probability to a global solution.
' 18 All variables must have both upper and lower bounds.
' 19 Variable bounds conflict in binary or alldifferent constraint.
' 20 Lower and upper bounds on variables allow no feasible solution.

' -----------------------------------------------------------------------------
' OpenSolver results, as given by OpenSolver.SolveStatus
' See also OpenSolver.SolveStatusString, which gives a slightly more detailed text summary
' and OpenSolver.SolveStatusComment, for any detailed comment on an infeasible problem
Enum OpenSolverResult
   Pending = -4  ' Added by us - used for solvers that asynchronously and are yet to run
   AbortedThruUserAction = -3   ' Used to indicate that a non-linearity check was made (losing the solution)
   ErrorOccurred = -2    ' Added by us - used in RunOpenSolver to indicate an error occured and has been reported to the user
   Unsolved = -1        ' Added by us - indicates a model not yet solved
   Optimal = 0
   Unbounded = 4        ' "The Objective Cell values do not converge" => unbounded
   Infeasible = 5
   LimitedSubOptimal = 10    ' CBC stopped before finding an optimal/feasible/integer solution because of CBC errors or time/iteration limits
   NotLinear = 7 ' Report non-linearity so that it can be picked up in silent mode
   ' ErrorInTargetOrConstraint = 9  We throw an error instead
   ' ErrorInModel = 13 We throw an error instead
   ' IntegerOptimal = 14 We just return Optimal
End Enum

' OpenSolver Solver Types
Enum OpenSolver_SolverType
    Unknown = -1
    Linear = 1
    NonLinear = 2
End Enum

Enum OpenSolver_ModelStatus
    Unitialized = 0
    Built = 1
End Enum

' The type of value in the LHS/RHS of a constraint
Enum SolverInputType
    SingleCellRange = 1 ' Valid for a LHS and a RHS
    MultiCellRange = 2  ' Valid for a LHS and a RHS
    Formula = 3         ' Valid for a RHS only
    constant = 4        ' Valid for a RHS only
End Enum

' Solver's different types of constraints
Public Enum RelationConsts
    [_First] = 1
    RelationLE = 1
    RelationEQ = 2
    RelationGE = 3
    RelationINT = 4
    RelationBIN = 5
    RelationAllDiff = 6
    [_Last] = 6
End Enum

Public Enum ObjectiveSenseType
   [_First] = 0
   UnknownObjectiveSense = 0
   MaximiseObjective = 1
   MinimiseObjective = 2
   TargetObjective = 3
   [_Last] = 3
End Enum

Public Enum VariableType
    VarContinuous = 0
    VarInteger = 1
    VarBinary = 2
End Enum

Public Enum OpenSolver_FileType
    LP = 1
    AMPL = 2
    NL = 3
End Enum

Public Enum OpenSolver_ModelType
    None = 1
    Diff = 2
    Parsed = 3
End Enum


#If Mac Then
    Public Const ExecExtension = ""
    Public Const ScriptExtension = ".sh"
    Public Const NBSP = 202 ' ascii char code for non-breaking space on Mac
    Public Const ScriptSeparator = " ; "
#Else
    Public Const ExecExtension = ".exe"
    Public Const ScriptExtension = ".bat"
    Public Const NBSP = 160 ' ascii char code for non-breaking space on Windows
    Public Const ScriptSeparator = " & "
#End If

Function ObjectiveSenseStringToEnum(ByVal sense As String) As ObjectiveSenseType
1         Select Case sense
          Case "min", "minimize", "minimise": ObjectiveSenseStringToEnum = MinimiseObjective
2         Case "max", "maximize", "maximise": ObjectiveSenseStringToEnum = MaximiseObjective
3         End Select
End Function

Function RelationStringToEnum(ByVal rel As String) As RelationConsts
1         Select Case rel
          Case "<", "<=", "=<":                                                     RelationStringToEnum = RelationLE
2         Case "=", "'=":                                                           RelationStringToEnum = RelationEQ
3         Case ">", ">=", "=>":                                                     RelationStringToEnum = RelationGE
4         Case "integer", "int", "i", "integers", "generals", "general", "gen":     RelationStringToEnum = RelationINT
5         Case "binary", "bin", "b", "binaries":                                    RelationStringToEnum = RelationBIN
6         Case "alldiff":                                                           RelationStringToEnum = RelationAllDiff
7         Case Else
8             RaiseGeneralError "Unknown relation code: " & rel
9         End Select
End Function

Function RelationEnumToString(rel As RelationConsts) As String
1         Select Case rel
          Case RelationLE:      RelationEnumToString = "<="
2         Case RelationEQ:      RelationEnumToString = "="
3         Case RelationGE:      RelationEnumToString = ">="
4         Case RelationINT:     RelationEnumToString = "int"
5         Case RelationBIN:     RelationEnumToString = "bin"
6         Case RelationAllDiff: RelationEnumToString = "alldiff"
7         End Select
End Function

Function SolverRelationAsUnicodeChar(rel As RelationConsts) As String
1         Select Case rel
          Case RelationGE: SolverRelationAsUnicodeChar = ChrW(&H2265) ' ">" gg
2         Case RelationEQ: SolverRelationAsUnicodeChar = "="
3         Case RelationLE: SolverRelationAsUnicodeChar = ChrW(&H2264) ' "<"
4         Case Else:       SolverRelationAsUnicodeChar = "(unknown)"
5         End Select
End Function

Function ReverseRelation(rel As Long) As Long
1         ReverseRelation 4 - rel
End Function

Function RelationHasRHS(rel As RelationConsts) As Boolean
1         Select Case rel
          Case RelationLE, RelationEQ, RelationGE:        RelationHasRHS = True
2         Case RelationINT, RelationBIN, RelationAllDiff: RelationHasRHS = False
3         End Select
End Function

