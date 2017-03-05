VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COpenSolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

' Model information
Public sheet As Worksheet

' Transient options
Public MinimiseUserInteraction As Boolean
Public SolveRelaxation As Boolean

' Permanent options
Dim pSolverParameters As Dictionary
Public ShowIterationResults As Boolean
Public AssumeNonNegativeVars As Boolean

' Variable information
Public AdjustableCells As Range    ' All decision variables
Public BinaryCellsRange As Range   ' All binary variables
Public IntegerCellsRange As Range  ' All integer variables (excluding binary)
Public VarLowerBounds As Dictionary
Public VarUpperBounds As Dictionary
Private pVarName() As String
Private pVarInitialValue() As Double
Public VarNameToIndex As Collection
Private pVarCategory() As VariableType  ' Name to avoid clash with VarType() function
Public InitialSolutionIsValid As Boolean  ' True if we can warmstart with the starting solution

' Constraint information
Dim pLHSType() As SolverInputType
Dim pRHSType() As SolverInputType
Dim pLHSRange() As Range
Dim pRHSRange() As Range
Dim pRHSFormula() As String
Dim pRHSConstant() As Double
Dim pRelation() As RelationConsts ' For each row, we remember its relation (=, <=, >=)
Dim pConstraintSummary() As String

' Objective information
Public ObjRange As Range
Public ObjectiveSense As ObjectiveSenseType
Public ObjectiveTargetValue As Double

' Details for the chosen solver
Dim pSolver As ISolver
Public ExternalSolverPathName As String 'Path to solver
Public LogFilePathName As String
Public SolutionFilePathName As String

' Details we need for non-linear solver
Dim pRowSetsBound() As Boolean

' Details for sensitivity analysis
Public bGetDuals As Boolean ' Whether to get duals
Public rConstraintList As Range
Public DualsOnNewSheet As Boolean
Public DualsOnSameSheet As Boolean

' The model builders
Public DiffModel As CModelDiff
Public ParsedModel As CModelParsed

' Build results
Public ModelStatus As OpenSolver_ModelStatus
Public NumVars As Long
Public NumIntVars As Long
Public NumBinVars As Long
Public NumDiscreteVars As Long
Public NumConstraints As Long  ' Number of Solver-constraints in the model
Public NumRows As Long         ' Number of constraints in the A matrix

' Model information filled in by specific model builder
Dim pSparseA() As CIndexedCoeffs            ' This is the A matrix for our Solver model
Dim pRHS() As Double                        ' The RHS values for each row in the LP
Public CostCoeffs As CIndexedCoeffs            ' The sparse cost vector
Public ObjectiveFunctionConstant As Double  ' The constant offset in the objective function


Dim ConstraintToRow() As Long   ' Maps solver constraint number to row
Dim pRowToConstraint() As Long  ' Maps row number to original solver constraint number
Dim RowCount() As Long          ' Number of rows for each solver constraint

' SOLUTION INFORMATION
' Constraints
Dim pConFinalValue() As Double  ' Final value of each constraint (RHScell - LHScell)
Dim pConShadowPrice() As Double ' Shadow price for each constraint
Dim pConIncrease() As Double    ' Allowable increase for each constraint
Dim pConDecrease() As Double    ' Allowable decrease for each constraint
' Variables
Dim pVarCellName() As String    ' Name of Excel cell for each variable (must be inferred from model/solution files)
Dim pVarFinalValue() As Double  ' Final value of each variable
Dim pVarReducedCost() As Double ' Reduced cost for each variable
Dim pVarIncrease() As Double    ' Allowable increase for each variable
Dim pVarDecrease() As Double    ' Allowable decrease for each variable

' Our result as an enum and a string
Public SolveStatus As OpenSolverResult
Public SolveStatusString As String    ' Eg Infeasible, Opimal etc
Public SolveStatusComment As String ' Eg constraint 1 cannot be satisified
Public SolutionWasLoaded As Boolean ' True if we loaded a solution into the spreadsheet

' The following are calculated differently; TimeTakenToBuildModel fails if we cross midnight!
Dim TimeTakenToBuildModel As Single, SecondsTakenToBuildModel As Long

' Support for doing a quick-solve where we only modify the RHS
Dim QuickSolve As CQuickSolve

' ============================================================================
' Array accessors
Public Property Get VarFinalValue(Index As Long) As Double
1         VarFinalValue = pVarFinalValue(Index)
End Property
Public Property Let VarFinalValue(Index As Long, value As Double)
1         If Index > UBound(pVarFinalValue) Then ReDim Preserve pVarFinalValue(Index)
2         pVarFinalValue(Index) = value
End Property

Public Property Get ConFinalValue(Index As Long) As Double
1         ConFinalValue = pConFinalValue(Index)
End Property
Public Property Let ConFinalValue(Index As Long, value As Double)
1         If Index > UBound(pConFinalValue) Then ReDim Preserve pConFinalValue(Index)
2         pConFinalValue(Index) = value
End Property

Public Property Get ConShadowPrice(Index As Long) As Double
1         ConShadowPrice = pConShadowPrice(Index)
End Property
Public Property Let ConShadowPrice(Index As Long, value As Double)
1         If Index > UBound(pConShadowPrice) Then ReDim Preserve pConShadowPrice(Index)
2         pConShadowPrice(Index) = value
End Property

Public Property Get VarReducedCost(Index As Long) As Double
1         VarReducedCost = pVarReducedCost(Index)
End Property
Public Property Let VarReducedCost(Index As Long, value As Double)
1         If Index > UBound(pVarReducedCost) Then ReDim Preserve pVarReducedCost(Index)
2         pVarReducedCost(Index) = value
End Property

Public Property Get VarCellName(Index As Long) As String
1         VarCellName = pVarCellName(Index)
End Property
Public Property Let VarCellName(Index As Long, value As String)
1         If Index > UBound(pVarCellName) Then ReDim Preserve pVarCellName(Index)
2         pVarCellName(Index) = value
End Property

Public Property Get ConIncrease(Index As Long) As Double
1         ConIncrease = pConIncrease(Index)
End Property
Public Property Let ConIncrease(Index As Long, value As Double)
1         If Index > UBound(pConIncrease) Then ReDim Preserve pConIncrease(Index)
2         pConIncrease(Index) = value
End Property

Public Property Get ConDecrease(Index As Long) As Double
1         ConDecrease = pConDecrease(Index)
End Property
Public Property Let ConDecrease(Index As Long, value As Double)
1         If Index > UBound(pConDecrease) Then ReDim Preserve pConDecrease(Index)
2         pConDecrease(Index) = value
End Property

Public Property Get VarIncrease(Index As Long) As Double
1         VarIncrease = pVarIncrease(Index)
End Property
Public Property Let VarIncrease(Index As Long, value As Double)
1         If Index > UBound(pVarIncrease) Then ReDim Preserve pVarIncrease(Index)
2         pVarIncrease(Index) = value
End Property

Public Property Get VarDecrease(Index As Long) As Double
1         VarDecrease = pVarDecrease(Index)
End Property
Public Property Let VarDecrease(Index As Long, value As Double)
1         If Index > UBound(pVarDecrease) Then ReDim Preserve pVarDecrease(Index)
2         pVarDecrease(Index) = value
End Property

Public Property Get VarName(Index As Long) As String
1         VarName = pVarName(Index)
End Property
Public Property Let VarName(Index As Long, value As String)
1               pVarName(Index) = value
End Property

Public Property Get VarInitialValue(Index As Long) As Double
1         VarInitialValue = pVarInitialValue(Index)
End Property
Public Property Let VarInitialValue(Index As Long, value As Double)
1               pVarInitialValue(Index) = value
End Property

Public Property Get VarCategory(Index As Long) As VariableType
1               VarCategory = pVarCategory(Index)
End Property
Public Property Let VarCategory(Index As Long, value As VariableType)
1               pVarCategory(Index) = value
End Property

Public Property Get LHSType(Index As Long) As SolverInputType
1               LHSType = pLHSType(Index)
End Property
Public Property Let LHSType(Index As Long, value As SolverInputType)
1               pLHSType(Index) = value
End Property

Public Property Get RHSType(Index As Long) As SolverInputType
1               RHSType = pRHSType(Index)
End Property
Public Property Let RHSType(Index As Long, value As SolverInputType)
1               pRHSType(Index) = value
End Property

Public Property Get LHSRange(Index As Long) As Range
1               Set LHSRange = pLHSRange(Index)
End Property
Public Property Set LHSRange(Index As Long, value As Range)
1               Set pLHSRange(Index) = value
End Property

Public Property Get RHSRange(Index As Long) As Range
1               Set RHSRange = pRHSRange(Index)
End Property
Public Property Set RHSRange(Index As Long, value As Range)
1               Set pRHSRange(Index) = value
End Property

Public Property Get RHSFormula(Index As Long) As String
1               RHSFormula = pRHSFormula(Index)
End Property
Public Property Let RHSFormula(Index As Long, value As String)
1               pRHSFormula(Index) = value
End Property

Public Property Get RHSConstant(Index As Long) As Double
1               RHSConstant = pRHSConstant(Index)
End Property
Public Property Let RHSConstant(Index As Long, value As Double)
1               pRHSConstant(Index) = value
End Property

Public Property Get Relation(Index As Long) As RelationConsts
1               Relation = pRelation(Index)
End Property
Public Property Let Relation(Index As Long, value As RelationConsts)
1               pRelation(Index) = value
End Property

Public Property Get RowSetsBound(Index As Long) As Boolean
1               RowSetsBound = pRowSetsBound(Index)
End Property
Public Property Let RowSetsBound(Index As Long, value As Boolean)
1               pRowSetsBound(Index) = value
End Property

Public Property Get ConstraintSummary(Index As Long) As String
1               ConstraintSummary = pConstraintSummary(Index)
End Property
Public Property Let ConstraintSummary(Index As Long, value As String)
1               pConstraintSummary(Index) = value
End Property

Public Property Set Solver(NewSolver As ISolver)
1               Set pSolver = NewSolver
End Property
Public Property Get Solver() As ISolver
1               Set Solver = pSolver
End Property

Public Property Set SolverParameters(NewSolverParameters As Dictionary)
1               Set pSolverParameters = NewSolverParameters
End Property
Public Property Get SolverParameters() As Dictionary
1               Set SolverParameters = pSolverParameters
End Property

Public Property Get RHS(Index As Long) As Double
1               RHS = pRHS(Index)
End Property
Public Property Let RHS(Index As Long, value As Double)
1               pRHS(Index) = value
End Property

Public Property Set SparseA(Index As Long, NewSparseA As CIndexedCoeffs)
1               Set pSparseA(Index) = NewSparseA
End Property
Public Property Get SparseA(Index As Long) As CIndexedCoeffs
1               Set SparseA = pSparseA(Index)
End Property

Public Property Get RowToConstraint(Index As Long) As Long
1               RowToConstraint = pRowToConstraint(Index)
End Property
Public Property Let RowToConstraint(Index As Long, value As Long)
1               pRowToConstraint(Index) = value
End Property
' End of getters and setters
'===============================================================================

Function GetConstraintInstance(ByVal row As Long, ByVal constraint As Long) As Long
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         GetConstraintInstance = row - ConstraintToRow(constraint) + 1

ExitFunction:
4         If RaiseError Then RethrowError
5         Exit Function

ErrorHandler:
6         If Not ReportError("COpenSolver", "GetConstraintInstance") Then Resume
7         RaiseError = True
8         GoTo ExitFunction
End Function

Function CheckConstraintValuesAreNumeric(constraint As Long, LHSValues As Variant, RHSValues As Variant, ErrorMsg As String, AllowNonNumericIfNotConstant As Boolean) As Boolean
          ' Check that the LHS and RHS values are doubles (not vbError, or vbString, or ??)
          ' vbEmpty (blank) cell values are permitted as VBA quietly turns them into 0's in expressions
1         CheckConstraintValuesAreNumeric = False
          Dim i As Long, j As Long, instance As Long
          Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String, cell As Range
          
          ' Check LHS first; this is always an array
2         instance = 0
3         For i = 1 To UBound(LHSValues, 1)
4             For j = 1 To UBound(LHSValues, 2)
5                 instance = instance + 1
6                 Select Case VarType(LHSValues(i, j))
                      Case vbDouble:
7                     Case vbEmpty:
                          ' This is an empty cell; VBA quietly turns this into 0 when used in an expression. We need no special handling of this, and don't report an error.
8                     Case vbError:
9                         GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
10                        ErrorMsg = "The cell '" & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "' in constraint '" & ConstraintSummary(constraint) & "' has a divide by zero error or has a problem with the function used ( eg #DIV/0! or #VALUE!)."
11                    Case Else:
                          ' It might be vbError, or vbString, or vbEmpty
12                        GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
13                        ErrorMsg = "The cell '" & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "' in constraint '" & ConstraintSummary(constraint) & "' has a value which is not numeric."
14                        On Error Resume Next
                          ' Try to show the incorrect value; just cary on with the old ErrorMsg if this gives an error
15                        ErrorMsg = "The cell '" & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "' in constraint '" & ConstraintSummary(constraint) & "' has the value '" & LHSValues(i, j) & "' which is not numeric."
16                        On Error GoTo 0
17                End Select

18                If Len(ErrorMsg) > 0 Then
                      ' We have an invalid value
19                    If AllowNonNumericIfNotConstant Then
                          ' Allow this if the cell depends on the variable cells - it might be valid later
20                        Set cell = GetOneCellInRange(LHSCellRange, instance)
21                        If Not ProperPrecedents(cell) Is Nothing Then
22                            If Not Intersect(ProperPrecedents(cell), AdjustableCells) Is Nothing Then
23                                GoTo NextCellLHS
24                            End If
25                        End If
26                    End If
                      ' Else the error is real, exit and report it
27                    GoTo ExitError
28                End If
NextCellLHS:
29                ErrorMsg = vbNullString
30            Next j
31        Next i
          
32        Select Case RHSType(constraint)
              Case SolverInputType.SingleCellRange, SolverInputType.Formula, SolverInputType.constant
33                Select Case VarType(RHSValues)
                      Case vbDouble
34                    Case vbEmpty:
                          ' This is an empty cell; VBA quietly turns this into 0 when used in an expression. We need no special handling of this, and don't report an error.
35                    Case vbError:
36                        GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
37                        ErrorMsg = "The right hand side in constraint '" & ConstraintSummary(constraint) & "' has a divide by zero error or has a problem with the function used (eg #DIV/0! or #VALUE!)."
38                    Case Else:
                          ' It might be vbError, or vbString, or anything else?
39                        GetConstraintInstanceData constraint, 1, LHSCellRange, RHSCellRange, RHSstring
40                        ErrorMsg = "The right hand side in constraint '" & ConstraintSummary(constraint) & "' has a value which is not a single numeric value."
                          ' Try to show the incorrect value; just cary on with the old ErrorMsg if this gives an error
41                        On Error Resume Next
42                        ErrorMsg = "The right hand side in constraint '" & ConstraintSummary(constraint) & "' has the value '" & RHSValues & "' which is not numeric."
43                        On Error GoTo 0
44                End Select

45                If Len(ErrorMsg) > 0 Then
                      ' We have an invalid value
46                    If AllowNonNumericIfNotConstant Then
                          ' Allow this if the cell depends on the variable cells - it might be valid later
47                        Set cell = RHSCellRange
48                        If Not ProperPrecedents(cell) Is Nothing Then
49                            If Not Intersect(ProperPrecedents(cell), AdjustableCells) Is Nothing Then
50                                GoTo NextCellSingleRHS
51                            End If
52                        End If
53                    End If
                      ' Else the error is real, exit and report it
54                    GoTo ExitError
55                End If
NextCellSingleRHS:
56            Case SolverInputType.MultiCellRange
57                instance = 0
58                For i = 1 To UBound(RHSValues, 1)
59                    For j = 1 To UBound(RHSValues, 2)
60                        instance = instance + 1
61                        Select Case VarType(RHSValues(i, j))
                              Case vbDouble:
62                            Case vbEmpty:
                                  ' This is an empty cell; VBA quietly turns this into 0 when used in an expression. We need no special handling of this, and don't report an error.
63                            Case vbError:
64                                GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
65                                ErrorMsg = "The right hand side in constraint '" & ConstraintSummary(constraint) & "' has a divide by zero error or has a problem with the function used ( eg #DIV/0! or #VALUE!)."
66                                GoTo ExitError
67                            Case Else:
                                  ' It might be vbError, or vbString, or vbEmpty
68                                GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
69                                ErrorMsg = "The cell '" & RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "' in constraint '" & ConstraintSummary(constraint) & "' has a value which is not numeric."
70                                On Error Resume Next
71                                ErrorMsg = "The cell '" & RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "' in constraint '" & ConstraintSummary(constraint) & "' has the value '" & RHSValues(i, j) & "' which is not numeric."
72                                On Error GoTo 0
73                        End Select

74                        If Len(ErrorMsg) > 0 Then
                              ' We have an invalid value
75                            If AllowNonNumericIfNotConstant Then
                                  ' Allow this if the cell depends on the variable cells - it might be valid later
76                                Set cell = GetOneCellInRange(RHSCellRange, instance)
77                                If Not ProperPrecedents(cell) Is Nothing Then
78                                    If Not Intersect(ProperPrecedents(cell), AdjustableCells) Is Nothing Then
79                                        GoTo NextCellMultiRHS
80                                    End If
81                                End If
82                            End If
                              ' Else the error is real, exit and report it
83                            GoTo ExitError
84                        End If
NextCellMultiRHS:
85                        ErrorMsg = vbNullString
86                    Next j
87                Next i
88        End Select
89        CheckConstraintValuesAreNumeric = True
90        Exit Function
          
ExitError:
91        ErrorMsg = ErrorMsg & vbNewLine & vbNewLine & _
                     IIf(AllowNonNumericIfNotConstant, "This can happen if you have made an error when constructing the model.", _
                                                       "This can happen if your model is nonlinear, or if you have made an error when constructing the model.") & _
                     " Please fix this and try again."
End Function

' Get the current values off the spreadsheet for the given constraint. LHS is always returned as a 2D array. RHS may be a single or array value
' Options:
' - ValidateNumeric: validates that all values are numeric
' - AllowNonNumericIfNotConstant: doesn't throw an error if the non-numeric value depends on the variable cells
Sub GetCurrentConstraintValues(constraint As Long, LHSValues As Variant, RHSValues As Variant, Optional ValidateNumeric As Boolean = True, Optional AllowNonNumericIfNotConstant As Boolean = False)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If LHSType(constraint) = SolverInputType.SingleCellRange Then
              ' Create a 1x1 array
4             LHSValues = Create1x1Array(LHSRange(constraint).Value2)
5         Else ' LHSType(constraint) = SolverInputType.MultiCellRange
              ' Copy all the values into a variant, which will then be an
              ' array of variants; see http://support.microsoft.com/kb/213798
6             LHSValues = LHSRange(constraint).Value2
7         End If

8         If RHSType(constraint) = SolverInputType.SingleCellRange Then
              ' A variant containing just a single double
9             RHSValues = RHSRange(constraint).Value2
10        ElseIf RHSType(constraint) = SolverInputType.MultiCellRange Then
              ' Copy all the values into a variant, which will then be an
              ' array of variants
11            RHSValues = RHSRange(constraint).Value2
12        ElseIf RHSType(constraint) = SolverInputType.Formula Then
              ' Will not throw an error, but may put an error value in RHSValues
13            RHSValues = sheet.Evaluate("=" & RHSFormula(constraint))
14        Else  ' RHSType(constraint) = SolverInputType.Constant
15            RHSValues = RHSConstant(constraint)
16        End If

17        If ValidateNumeric Then
              Dim ErrorMsg As String
18            If Not CheckConstraintValuesAreNumeric(constraint, LHSValues, RHSValues, ErrorMsg, _
                                                     AllowNonNumericIfNotConstant) Then
19                RaiseUserError ErrorMsg
20            End If
21        End If

ExitSub:
22        If RaiseError Then RethrowError
23        Exit Sub

ErrorHandler:
24        If Not ReportError("COpenSolver", "GetCurrentConstraintValues") Then Resume
25        RaiseError = True
26        GoTo ExitSub
End Sub

Function GetCurrentObjectiveValue(Optional ValidateNumeric As Boolean = True, Optional AllowNonNumericIfNotConstant As Boolean = False) As Variant
                Dim RaiseError As Boolean
1               RaiseError = False
2               On Error GoTo ErrorHandler
                
3               If ObjRange Is Nothing Then
4                   GetCurrentObjectiveValue = 0
5                   GoTo ExitFunction
6               End If
                    
                Dim ObjValue As Variant
7               ObjValue = ObjRange.Value2
                
8               If ValidateNumeric Then
                    Dim ErrorMsg As String
                    ' Check if objective has a value that is not a number
9                   If VarType(ObjValue) <> vbDouble Then
10                      If VarType(ObjValue) = vbError Then
11                          ErrorMsg = "The objective cell appears to contain an error. This could have occurred if there is a divide by zero error or if you have used the wrong function (eg #DIV/0! or #VALUE!)."
12                      Else
13                          ErrorMsg = "The objective cell does not appear to contain a numeric value."
14                      End If
15                  End If
                    
16                  If Len(ErrorMsg) > 0 Then
17                      If AllowNonNumericIfNotConstant Then
18                          If Not ProperPrecedents(ObjRange) Is Nothing Then
19                              If Not Intersect(ProperPrecedents(ObjRange), AdjustableCells) Is Nothing Then
20                                  GoTo DontRaise
21                              End If
22                          End If
23                      End If
24                      ErrorMsg = _
                            ErrorMsg & vbNewLine & vbNewLine & _
                            IIf(AllowNonNumericIfNotConstant, "This can happen if you have made an error when constructing the model.", _
                                                              "This can happen if your model is nonlinear, or if you have made an error when constructing the model.") & _
                            " Please fix this and try again."
25                      RaiseUserError ErrorMsg
26                  End If
27              End If
DontRaise:
28              GetCurrentObjectiveValue = ObjValue

ExitFunction:
29              If RaiseError Then RethrowError
30              Exit Function

ErrorHandler:
31              If Not ReportError("COpenSolver", "GetCurrentObjectiveValue") Then Resume
32              RaiseError = True
33              GoTo ExitFunction
End Function


' Get the information required to display a single instance of a constraint to the user
Sub GetConstraintInstanceData(constraint As Long, instance As Long, LHSCellRange As Range, RHSCellRange As Range, RHSstring As String)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         Select Case LHSType(constraint)
              Case SolverInputType.SingleCellRange
4                 Set LHSCellRange = LHSRange(constraint)
5             Case SolverInputType.MultiCellRange
6                 Set LHSCellRange = GetOneCellInRange(LHSRange(constraint), instance)
7             Case SolverInputType.Formula
8                 Debug.Assert False ' should never happen
9             Case SolverInputType.constant
10                Debug.Assert False ' should never happen
11        End Select
12        Select Case RHSType(constraint)
              Case SolverInputType.SingleCellRange
13                Set RHSCellRange = RHSRange(constraint)
14                RHSstring = vbNullString
15            Case SolverInputType.MultiCellRange
16                Set RHSCellRange = GetOneCellInRange(RHSRange(constraint), instance)
17                RHSstring = vbNullString
18            Case SolverInputType.Formula, SolverInputType.constant
19                Set RHSCellRange = Nothing
20                RHSstring = RHSFormula(constraint)
21        End Select

ExitSub:
22        If RaiseError Then RethrowError
23        Exit Sub

ErrorHandler:
24        If Not ReportError("COpenSolver", "GetConstraintInstanceData") Then Resume
25        RaiseError = True
26        GoTo ExitSub
End Sub

Sub CheckCurrentConstraintFeasible(LHSValues As Variant, RHSValues As Variant, rel As RelationConsts, constraint As Long)
                ' Check if the current solution is valid for this constraint
1               On Error GoTo Invalid
                
                Dim i As Long, j As Long, LHSValue As Double, RHSValue As Double, value As Double
2               For i = 1 To UBound(LHSValues, 1)
3                   For j = 1 To UBound(LHSValues, 2)
4                       LHSValue = LHSValues(i, j)
5                       If RHSType(constraint) <> SolverInputType.MultiCellRange Then
6                           RHSValue = RHSValues
7                       ElseIf UBound(LHSValues, 1) = UBound(RHSValues, 1) Then
8                           RHSValue = RHSValues(i, j)
9                       Else
10                          RHSValue = RHSValues(j, i)
11                      End If
                              
12                      value = RHSValue - LHSValue
13                      If (rel = RelationEQ And Abs(value) > EPSILON) Or _
                           (rel = RelationGE And value > EPSILON) Or _
                           (rel = RelationLE And value < -EPSILON) Then
14                          GoTo Invalid
15                      End If
16                  Next j
17              Next i
18              Exit Sub
Invalid:
19              InitialSolutionIsValid = False
End Sub

Sub CheckCurrentVariablesFeasible(ShouldSolveRelaxation As Boolean)
                Dim i As Long, RoundedValue As Double, bound As Double
1               For i = 1 To NumVars
                    ' Check value is integral if discrete
2                   If VarCategory(i) <> VarContinuous And Not ShouldSolveRelaxation Then
                        ' If the starting value is not integer, solution is invalid.
                        ' Otherwise round starting value to ensure feasibility
3                       RoundedValue = Round(VarInitialValue(i))
4                       If Abs(VarInitialValue(i) - RoundedValue) <= EPSILON Then
5                           VarInitialValue(i) = RoundedValue
6                       Else
7                           GoTo Invalid
8                       End If
9                   End If
                    
                    ' Check bounds - we have already invalidated anything that violates these in the constraint validation
                    ' However, we go through them again because we can now adjust variables that are only just infeasible
                        
                    ' Check lower bounds
10                  If VarLowerBounds.Exists(VarName(i)) Or VarCategory(i) = VarBinary Or AssumeNonNegativeVars Then
11                      If VarLowerBounds.Exists(VarName(i)) And VarCategory(i) <> VarBinary Then
12                          bound = VarLowerBounds.Item(VarName(i))
13                      Else
14                          bound = 0
15                      End If
16                      If VarInitialValue(i) >= bound - EPSILON Then
                            ' Make sure the bound is strictly satisfied
17                          If VarInitialValue(i) < bound Then
18                              VarInitialValue(i) = bound
19                          End If
20                      Else
21                          GoTo Invalid
22                      End If
23                  End If
                    
                    ' Check upper bounds
24                  If VarUpperBounds.Exists(VarName(i)) Or VarCategory(i) = VarBinary Then
25                      If VarCategory(i) = VarBinary Then
26                          bound = 1
27                      Else
28                          bound = VarUpperBounds.Item(VarName(i))
29                      End If
                        
30                      If VarInitialValue(i) <= bound + EPSILON Then
                            ' Make sure the bound is strictly satisfied
31                          If VarInitialValue(i) > bound Then
32                              VarInitialValue(i) = bound
33                          End If
34                      Else
35                          GoTo Invalid
36                      End If
37                  End If
38              Next i
39              Exit Sub
Invalid:
40              InitialSolutionIsValid = False
End Sub

Sub BuildModelFromSolverData(Optional LinearityOffset As Double = 0, _
                             Optional ShouldCheckLinearity As Boolean = False, _
                             Optional ShouldMinimiseUserInteraction As Boolean = False, _
                             Optional ShouldSolveRelaxation As Boolean = False, _
                             Optional modelSheet As Worksheet)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         SolveStatus = OpenSolverResult.Unsolved
4         SolveStatusString = "Unsolved"
5         SolveStatusComment = vbNullString
6         SolutionWasLoaded = False
7         ModelStatus = Unitialized

8         MinimiseUserInteraction = ShouldMinimiseUserInteraction

          ' Change to manual calculation mode
          Dim oldCalculationMode As Long
9         oldCalculationMode = Application.Calculation
10        Application.Calculation = xlCalculationManual

          Dim ScreenStatus As Boolean
11        ScreenStatus = Application.ScreenUpdating
12        Application.ScreenUpdating = False

13        Application.Cursor = xlWait

          Dim StartTimeOfBuild As Single, StartDateTimeOfBuild As Date
          Dim EndTimeOfBuild As Single, EndDateTimeOfBuild As Date
14        StartTimeOfBuild = Timer    ' Warning: This goes back to zero at midnight!
15        StartDateTimeOfBuild = Now()

16        Application.EnableCancelKey = xlErrorHandler
          
17        GetActiveSheetIfMissing modelSheet
18        Set sheet = modelSheet
          
          ' Find the solver
19        Set Solver = CreateSolver(GetChosenSolver(sheet))
20        LastUsedSolver = Solver.ShortName

          '--------------------------------------------------------------------------
          ' ADJUSTABLE CELLS SETUP
          
          ' Remove any overlap
21        Set AdjustableCells = GetDecisionVariablesNoOverlap(sheet)
          ' Handle any merged ranges
          Dim BadCell As Range
22        If Not CheckRangeContainsNoAmbiguousMergedCells(AdjustableCells, BadCell) Then
23            RaiseUserError "This model contains a decision variable " & BadCell.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & " that is inaccessible as it is within the merged range " & BadCell.MergeArea.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & ". Please redefine the decision variable cells, and try again."
24        End If
          ' Check we have any cells after all this
25        NumVars = AdjustableCells.Count
26        If NumVars < 1 Then
27            RaiseUserError "There are no decision variable cells in this model, and so the model cannot be solved."
28        End If

          ' Variable initialization
29        ReDim pVarName(1 To NumVars) As String
30        ReDim pVarInitialValue(1 To NumVars) As Double
31        ReDim pVarCategory(1 To NumVars) As VariableType
32        Set VarNameToIndex = New Collection
33        InitialSolutionIsValid = True
          
          Dim AdjCell As Range, i As Long
34        i = 1
35        For Each AdjCell In AdjustableCells
36            VarName(i) = GetCellName(AdjCell)
37            VarCategory(i) = VariableType.VarContinuous
38            VarNameToIndex.Add Item:=i, Key:=VarName(i)
              ' Try to get the starting value, if we can't the starting solution is invalid
39            On Error Resume Next
40            VarInitialValue(i) = AdjCell.Value2
41            If Err.Number <> 0 Then InitialSolutionIsValid = False
42            On Error GoTo ErrorHandler
43            i = i + 1
44        Next AdjCell
          
          ' Reset the cell values to catch protected sheets
45        TestCellsForWriting AdjustableCells
          
          ' Initial sheet calculation
46        If Not ForceCalculate("Warning: The initial worksheet calculation did not complete, and so the model may not be generated correctly. Would you like to try again?", MinimiseUserInteraction) Then
47            On Error GoTo ErrorHandler
48            RaiseGeneralError "Calculation failed. The model has not been built so cannot be solved."
49        End If

          '--------------------------------------------------------------------------
          ' OBJECTIVE SETUP
50        ObjectiveSense = GetObjectiveSense(sheet)
51        ObjectiveTargetValue = GetObjectiveTargetValue(sheet)
52        Set ObjRange = GetObjectiveFunctionCell(sheet, Validate:=True)
          ' Check that objective is not an non-numeric constant
53        GetCurrentObjectiveValue AllowNonNumericIfNotConstant:=True

          '--------------------------------------------------------------------------
          ' MODEL OPTIONS SETUP
54        AssumeNonNegativeVars = GetNonNegativity(sheet)
55        ShowIterationResults = GetShowSolverProgress(sheet)
          
          ' Collate all solver parameters (ours and user-defined)
56        Set SolverParameters = GetSolverParametersDict(Solver, sheet)

          ' Check for Solver's "Ignore Integer Constraints" option unless solving relaxation
57        If Not ShouldSolveRelaxation Then
              ' If we are in Excel 2010 or newer, and the user has set "Ignore Integer Constraints", then we report an error
58            If Val(Application.Version) >= 14 And GetIgnoreIntegerConstraints(sheet) Then
59                 If MinimiseUserInteraction Then
60                     RaiseUserError "You have the Solver option 'ignore integer constraints' turned on. Please turn this off in the Solver (not OpenSolver) options menu and then solve with OpenSolver again or choose OpenSolver's solve relaxed model option"
61                 ElseIf Not MsgBox("You have the 'Ignore integer constraints' option enabled in Solver, so OpenSolver cannot proceed with solving. You can either disable this option, or use the 'Solve Relaxation' option in the OpenSolver menu. Would you like OpenSolver to disable this option and proceed?", vbCritical + vbYesNo + vbDefaultButton1, "OpenSolver Warning") = vbNo Then
62                     SetIgnoreIntegerConstraints False, sheet
63                 Else
                       ' Already confirmed abort by user
64                     RaiseUserCancelledError
65                 End If
66            End If
67        End If
          
          '--------------------------------------------------------------------------
          ' CONSTRAINTS SETUP
68        NumConstraints = GetNumConstraints(sheet)
69        UpdateStatusBar "OpenSolver: Setting Up Problem... " & NumVars & " vars, " & NumConstraints & " Solver constraints", True
          'Set up arrays to be the correct size. We allow no constraints!
70        If NumConstraints > 0 Then
71            ReDim pLHSType(NumConstraints) As SolverInputType
72            ReDim pLHSRange(NumConstraints) As Range
73            ReDim pRelation(NumConstraints) As RelationConsts
74            ReDim pRHSType(NumConstraints) As SolverInputType
75            ReDim pRHSRange(NumConstraints) As Range
76            ReDim pRHSFormula(NumConstraints) As String
77            ReDim pRHSConstant(NumConstraints) As Double
78            ReDim pConstraintSummary(NumConstraints) As String
79            ReDim ConstraintToRow(NumConstraints) As Long
80            ReDim RowCount(NumConstraints) As Long
81        End If
82        Set VarLowerBounds = New Dictionary
83        Set VarUpperBounds = New Dictionary
84        ReDim pRowSetsBound(1)

85        NumRows = 0
          Dim constraint As Long, row As Long
86        row = 1
          ' Iterate through each solver constraint
87        For constraint = 1 To NumConstraints
88            UpdateStatusBar "OpenSolver: Setting Up Constraints... " & NumVars & " vars, " & constraint & "/" & NumConstraints & " Solver constraints"
                           
              ' LEFT HAND SIDE
              Dim rangeLHS As Range, sRefersToLHS As String
89            Set rangeLHS = GetConstraintLhs(constraint, sheet, Validate:=True, RefersTo:=sRefersToLHS)
              Dim LHSCount As Long
90            LHSCount = rangeLHS.Count
              
              ' RELATIONSHIP
              Dim rel As RelationConsts
91            rel = GetConstraintRel(constraint, sheet)
92            Relation(constraint) = rel
              
93            Select Case rel
              Case RelationINT, RelationBIN
                  ' Check that they haven't set as integer or binary something that is not a decision variable...
                  Dim intersection As Range
94                Set intersection = Intersect(rangeLHS, AdjustableCells)
95                If intersection Is Nothing Then
                      ' The cells are not all decision variables
96                    On Error GoTo ErrorHandler
97                    RaiseUserError "A range of cells are specified as bin or int that are not decision variables. OpenSolver does not support this."
98                End If
99                If intersection.Count = rangeLHS.Count Then
100                   If rel = RelationINT Then
101                       Set IntegerCellsRange = ProperUnion(IntegerCellsRange, rangeLHS)
102                   Else
103                       Set BinaryCellsRange = ProperUnion(BinaryCellsRange, rangeLHS)
104                   End If
105               Else
                      ' The cells are not all decision variables
106                   On Error GoTo ErrorHandler
107                   RaiseUserError "Unable to find a cell specified as bin or int in the decision variables."
108               End If
                  ' This constraint does not relate to any new rows, so mark it
                  ' with a dummy entry. row = same value as we get for the
                  ' next entry, so this refers to no rows.
109               ConstraintToRow(constraint) = row
              
              ' RIGHT HAND SIDE
110           Case RelationGE, RelationLE, RelationEQ
                  ' The constraint is a full equation with a RHS
                  Dim valRHS As Double, rangeRHS As Range, sRefersToRHS As String, RHSRefersToFormula As Boolean
111               Set rangeRHS = GetConstraintRhs(constraint, sRefersToRHS, valRHS, RHSRefersToFormula, sheet)
                  
112               ConstraintSummary(constraint) = StripWorksheetNameAndDollars(sRefersToLHS, sheet) & _
                                                  " " & RelationEnumToString(rel) & " " & _
                                                  StripWorksheetNameAndDollars(sRefersToRHS, sheet)
                                                  
                  Dim RHSCount As Long
113               If Not rangeRHS Is Nothing Then
114                   RHSCount = rangeRHS.Count
115               Else
116                   RHSCount = 1
117               End If
                  
                  ' Check we have a compatible constraint system.
118               If (LHSCount <> 1 And RHSCount <> 1) And (LHSCount <> RHSCount) Then
119                   On Error GoTo ErrorHandler
120                   RaiseUserError "The constraint '" & ConstraintSummary(constraint) & "' has a different cell count on the left and the right. The model cannot be built."
121               End If
                  
                  ' Store left hand side:
122               Set LHSRange(constraint) = rangeLHS
123               LHSType(constraint) = IIf(LHSCount = 1, SolverInputType.SingleCellRange, SolverInputType.MultiCellRange)
                  
                  ' Store right hand side:
124               If Not rangeRHS Is Nothing Then
125                   Set RHSRange(constraint) = rangeRHS
126                   RHSType(constraint) = IIf(RHSCount = 1, SolverInputType.SingleCellRange, SolverInputType.MultiCellRange)
127               ElseIf RHSRefersToFormula Then
                      ' Store the formula internally as a string
128                   RHSFormula(constraint) = sRefersToRHS
129                   RHSType(constraint) = SolverInputType.Formula
130               Else
                      ' Store a constant
131                   RHSFormula(constraint) = sRefersToRHS
132                   RHSType(constraint) = SolverInputType.constant
133                   RHSConstant(constraint) = valRHS
134               End If

                  ' Check that all constant values are numeric
                  ' We allow non-numeric values that might later change with the variable cells
                  Dim LHSValues As Variant
                  Dim RHSValues As Variant
135               GetCurrentConstraintValues constraint, LHSValues, RHSValues, AllowNonNumericIfNotConstant:=True
                  ' Check if the constraint is satisfied (unless we already know the solution is invalid)
136               If InitialSolutionIsValid Then
137                   CheckCurrentConstraintFeasible LHSValues, RHSValues, rel, constraint
138               End If
                  
                  ' Check if this constraint applies any bounds on decision variables
139               If LHSCount = 1 Then
                      ' Single LHS, single RHS
140                   ApplyBounds row, rel, rangeLHS, rangeRHS, valRHS, RHSRefersToFormula
141               Else
                      Dim CurRow As Long
142                   If RHSCount > 1 Then
                          ' multiple LHS, multiple RHS
143                       For CurRow = 1 To LHSCount
144                           ApplyBounds row + CurRow - 1, rel, rangeLHS(CurRow), rangeRHS(CurRow), valRHS, RHSRefersToFormula
145                       Next CurRow
146                   Else
                          ' multiple LHS, single RHS
147                       For CurRow = 1 To LHSCount
148                           ApplyBounds row + CurRow - 1, rel, rangeLHS(CurRow), rangeRHS, valRHS, RHSRefersToFormula
149                       Next CurRow
150                   End If
151               End If
                  
                  ' Store the row/constraint information
152               RowCount(constraint) = LHSCount
153               ConstraintToRow(constraint) = row
154               row = row + LHSCount
              
155           Case Else
                  ' Don't allow 'alldiff' from Excel 2010 or any other wierd things
156               On Error GoTo ErrorHandler
157               RaiseGeneralError "Unrecognised relationship for constraint " & ConstraintSummary(constraint)
158           End Select
159       Next constraint
160       NumRows = row - 1
          
161       If NumRows > 0 Then
              ' Make the row to constraint mapping
162           ReDim pRowToConstraint(NumRows) As Long
163           row = 0
164           For constraint = 1 To NumConstraints
165               For i = 1 To RowCount(constraint)
166                   RowToConstraint(row + i) = constraint
167               Next i
168               row = row + RowCount(constraint)
169           Next constraint
170       End If

          ' Binary trumps integer, so remove all overlap
171       Set IntegerCellsRange = SetDifference(IntegerCellsRange, BinaryCellsRange)
          
          Dim RoundedValue As Long
172       If Not IntegerCellsRange Is Nothing Then
173           Me.NumIntVars = IntegerCellsRange.Count
174           For Each AdjCell In IntegerCellsRange
175               VarCategory(VarNameToIndex(GetCellName(AdjCell))) = VarInteger
176           Next AdjCell
177       End If
178       If Not BinaryCellsRange Is Nothing Then
179           Me.NumBinVars = BinaryCellsRange.Count
180           For Each AdjCell In BinaryCellsRange
181               Me.VarCategory(VarNameToIndex(GetCellName(AdjCell))) = VarBinary
182           Next AdjCell
183       End If
184       Me.NumDiscreteVars = Me.NumIntVars + Me.NumBinVars
          
          ' Check whether variables are within bounds and discrete if needed
185       CheckCurrentVariablesFeasible ShouldSolveRelaxation

186       If NumRows > 0 Then
187           ReDim Preserve pRowSetsBound(NumRows)
188       End If
          
          'Set up extra model info depending on the solver
189       Select Case Solver.ModelType
          Case Diff
190           If NumRows > 0 Then
191               ReDim Preserve pSparseA(NumRows) As CIndexedCoeffs
192               ReDim pRHS(NumRows) As Double
193           End If
194           ReDim pCostCoeffs(NumVars) As Double
              
195           Set DiffModel = New CModelDiff
196           If Not DiffModel.ProcessSolverModel(Me, LinearityOffset, ShouldCheckLinearity) Then GoTo ExitSub
197       Case Parsed
              ' Parse spreadsheet and build model
198           Set ParsedModel = New CModelParsed
199           If Not ParsedModel.ProcessSolverModel(Me) Then GoTo ExitSub
200       End Select

201       ModelStatus = Built
202       EndTimeOfBuild = Timer
203       EndDateTimeOfBuild = Now()
204       SecondsTakenToBuildModel = DateDiff("s", StartDateTimeOfBuild, EndDateTimeOfBuild)
205       TimeTakenToBuildModel = EndTimeOfBuild - StartTimeOfBuild

ExitSub:
206       Application.Cursor = xlDefault
207       Application.StatusBar = False
208       Application.ScreenUpdating = ScreenStatus
209       Application.Calculation = oldCalculationMode
210       If RaiseError Then RethrowError
211       Exit Sub

ErrorHandler:
212       If Not ReportError("COpenSolver", "BuildModelFromSolverData") Then Resume
213       RaiseError = True
214       GoTo ExitSub
End Sub

Sub ApplyBounds(row As Long, rel As Long, LHS As Range, RHS As Range, RHSValue As Double, RHSRefersToFormula As Boolean)
          ' Determines whether the current LHS and RHS should be applied as a bound

          ' Excel 2007 requires the following conditions to be met for a user-supplied lower bound to exist
          '  The LHS range includes a decision variable directly (not via a formula referring to the decision var).
          '  The RHS is either: a constant, or: a (single or multicell) range that containing a constant or a formula. The RHS may not
          '  be specified as a formula (even one evaluating to a constant)
          ' Note: Excel 2010 seems a bit different in that the WHOLE LHS range must be decision variables. We adopt the 2007 behaviour
          ' Note: A constraint of the form 'range <= decision var' does NOT count as a lower bound; the decision var MUST be in the LHS
              
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If rel <> RelationGE And rel <> RelationLE Then GoTo ExitSub
           
          ' Following Solver, we don't apply formulae as lower bounds, even if constant
4         If RHSRefersToFormula Then GoTo ExitSub
          
          ' Check if LHS cell is a variable cell
5         If Intersect(LHS, AdjustableCells) Is Nothing Then GoTo ExitSub
          
6         If Not RHS Is Nothing Then
7             If IsNumeric(RHS.value) Then
8                 If Not RHS.HasFormula Then
                      ' The RHS range is a constant - apply a bound
9                     UpdateBound row, GetCellName(LHS), CDbl(RHS.value), rel
10                Else
                      ' The RHS range is a formula, check if this depends on the decision vars
                      ' Will raise an error is no precedents are found (e.g. if they are on another sheet
11                    If Intersect(RHS.Precedents, AdjustableCells) Is Nothing Then
12                        UpdateBound row, GetCellName(LHS), CDbl(RHS.value), rel
13                    End If
14                End If
15            End If
16        Else
              ' The RHS is a numeric constant - apply a bound
17            UpdateBound row, GetCellName(LHS), RHSValue, rel
18        End If
          
ExitSub:
19        If RaiseError Then RethrowError
20        Exit Sub

ErrorHandler:
          ' Ignore a 'No cells were found' precedents error
21        If Err.Number = 1004 Then
22            GoTo ExitSub
23        End If

24        If Not ReportError("COpenSolver", "ApplyBounds") Then Resume
25        RaiseError = True
26        GoTo ExitSub
End Sub

Sub UpdateBound(row As Long, VarName As String, bound As Double, rel As Long)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          ' Only GE constraints can provide lower bounds
3         If rel = RelationGE Then
              ' Remove existing lower bound on variable if our new one is better
4             If VarLowerBounds.Exists(VarName) Then
5                 VarLowerBounds.Item(VarName) = Max(bound, VarLowerBounds.Item(VarName))
6             Else
7                 VarLowerBounds.Add Item:=bound, Key:=VarName
8             End If
          ' Only LE constraints can provide upper bounds
9         ElseIf rel = RelationLE Then
              ' Remove existing upper bound on variable if our new one is better
10            If VarUpperBounds.Exists(VarName) Then
11                VarUpperBounds.Item(VarName) = Min(bound, VarUpperBounds.Item(VarName))
12            Else
13                VarUpperBounds.Add Item:=bound, Key:=VarName
14            End If
15        End If
          
          ' Track that this constraint row applies a bound
16        If UBound(pRowSetsBound) < row Then ReDim Preserve pRowSetsBound(row)
17        RowSetsBound(row) = True

ExitSub:
18        If RaiseError Then RethrowError
19        Exit Sub

ErrorHandler:
20        If Not ReportError("COpenSolver", "UpdateBound") Then Resume
21        RaiseError = True
22        GoTo ExitSub
End Sub

Sub ReportAnySolutionSubOptimality()
1               Select Case SolveStatus
                Case OpenSolverResult.Optimal                ' We are optimal
2               Case OpenSolverResult.Pending                ' The solve is scheduled
3               Case OpenSolverResult.NotLinear              ' Already shown non-linearity details
4               Case OpenSolverResult.AbortedThruUserAction  ' Already shown escape dialog
             
5               Case Else
6                   MsgBox "OpenSolver could not find an optimal solution, and reported:" & vbCrLf & _
                           SolveStatusString & vbCrLf & vbCrLf & _
                           IIf(Len(SolveStatusComment) > 0, SolveStatusComment & vbCrLf & vbCrLf, vbNullString) & _
                           IIf(SolutionWasLoaded, "The solution generated has been loaded into the spreadsheet.", "No solution was available to load into the spreadsheet.")
7               End Select
End Sub

Sub ClearQuickSolve()
1         Set QuickSolve = Nothing
End Sub

Function InitializeQuickSolve(ParamRange As Range) As Boolean
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         InitializeQuickSolve = False
4         If ModelStatus <> Built Then
5             RaiseUserError "QuickSolve Initialisation Failed: The model cannot be solved as it has not yet been built."
6         End If

7         Set QuickSolve = New CQuickSolve
8         If QuickSolve.AnalyseParameters(Me, ParamRange) Then
9             InitializeQuickSolve = True
10        Else
11            Set QuickSolve = Nothing
12        End If

ExitFunction:
13        If RaiseError Then RethrowError
14        Exit Function

ErrorHandler:
15        If Not ReportError("COpenSolver", "InitializeQuickSolve") Then Resume
16        RaiseError = True
17        GoTo ExitFunction
End Function

Sub DoQuickSolve(Optional SolveRelaxation As Boolean = False, Optional MinimiseUserInteraction As Boolean = False)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
3         If QuickSolve Is Nothing Then
4             RaiseUserError "Quick Solve has not been initialized. Please choose Intialize Quick Solve, and then try again."
5         Else
6             If QuickSolve.UpdateModelRHS(sheet, pRHS) Then
7                 SolveModel Me, SolveRelaxation, MinimiseUserInteraction
8             End If
9         End If

ExitSub:
10        If RaiseError Then RethrowError
11        Exit Sub

ErrorHandler:
12        If Not ReportError("COpenSolver", "DoQuickSolve") Then Resume
13        RaiseError = True
14        GoTo ExitSub
End Sub

Public Sub PrepareForSolution()
1         ReDim pVarFinalValue(NumVars) As Double
2         ReDim pVarCellName(NumVars) As String
3         If bGetDuals Then
4             ReDim pConFinalValue(NumRows) As Double
5             ReDim pConShadowPrice(NumRows) As Double
6             ReDim pConIncrease(NumRows) As Double
7             ReDim pConDecrease(NumRows) As Double
8             ReDim pVarReducedCost(NumVars) As Double
9             ReDim pVarIncrease(NumVars) As Double
10            ReDim pVarDecrease(NumVars) As Double
11        End If
End Sub

Public Sub LoadResultsToSheet()
1        If SolutionWasLoaded Then
2             AdjustableCells.Value2 = 0
              Dim i As Long
3             For i = 1 To NumVars
4                 If Len(VarCellName(i)) <> 0 Then
5                     AdjustableCells.Worksheet.Range(VarCellName(i)).Value2 = VarFinalValue(i)
6                 End If
7             Next i
8         End If
End Sub
