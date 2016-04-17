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
Public VarLowerBounds As Collection
Public VarUpperBounds As Collection
Private pVarName() As String
Public VarNameToIndex As Collection
Private pVarCategory() As VariableType  ' Name to avoid clash with VarType() function

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
Dim pCostCoeffs() As Double                 ' The cost vector
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
793       VarFinalValue = pVarFinalValue(Index)
End Property
Public Property Let VarFinalValue(Index As Long, value As Double)
794       If Index > UBound(pVarFinalValue) Then ReDim Preserve pVarFinalValue(Index)
795       pVarFinalValue(Index) = value
End Property

Public Property Get ConFinalValue(Index As Long) As Double
796       ConFinalValue = pConFinalValue(Index)
End Property
Public Property Let ConFinalValue(Index As Long, value As Double)
797       If Index > UBound(pConFinalValue) Then ReDim Preserve pConFinalValue(Index)
798       pConFinalValue(Index) = value
End Property

Public Property Get ConShadowPrice(Index As Long) As Double
799       ConShadowPrice = pConShadowPrice(Index)
End Property
Public Property Let ConShadowPrice(Index As Long, value As Double)
800       If Index > UBound(pConShadowPrice) Then ReDim Preserve pConShadowPrice(Index)
801       pConShadowPrice(Index) = value
End Property

Public Property Get VarReducedCost(Index As Long) As Double
802       VarReducedCost = pVarReducedCost(Index)
End Property
Public Property Let VarReducedCost(Index As Long, value As Double)
803       If Index > UBound(pVarReducedCost) Then ReDim Preserve pVarReducedCost(Index)
804       pVarReducedCost(Index) = value
End Property

Public Property Get VarCellName(Index As Long) As String
805       VarCellName = pVarCellName(Index)
End Property
Public Property Let VarCellName(Index As Long, value As String)
806       If Index > UBound(pVarCellName) Then ReDim Preserve pVarCellName(Index)
807       pVarCellName(Index) = value
End Property

Public Property Get ConIncrease(Index As Long) As Double
808       ConIncrease = pConIncrease(Index)
End Property
Public Property Let ConIncrease(Index As Long, value As Double)
809       If Index > UBound(pConIncrease) Then ReDim Preserve pConIncrease(Index)
810       pConIncrease(Index) = value
End Property

Public Property Get ConDecrease(Index As Long) As Double
811       ConDecrease = pConDecrease(Index)
End Property
Public Property Let ConDecrease(Index As Long, value As Double)
812       If Index > UBound(pConDecrease) Then ReDim Preserve pConDecrease(Index)
813       pConDecrease(Index) = value
End Property

Public Property Get VarIncrease(Index As Long) As Double
814       VarIncrease = pVarIncrease(Index)
End Property
Public Property Let VarIncrease(Index As Long, value As Double)
815       If Index > UBound(pVarIncrease) Then ReDim Preserve pVarIncrease(Index)
816       pVarIncrease(Index) = value
End Property

Public Property Get VarDecrease(Index As Long) As Double
817       VarDecrease = pVarDecrease(Index)
End Property
Public Property Let VarDecrease(Index As Long, value As Double)
818       If Index > UBound(pVarDecrease) Then ReDim Preserve pVarDecrease(Index)
819       pVarDecrease(Index) = value
End Property

Public Property Get VarName(Index As Long) As String
820       VarName = pVarName(Index)
End Property
Public Property Let VarName(Index As Long, value As String)
          pVarName(Index) = value
End Property

Public Property Get VarCategory(Index As Long) As VariableType
          VarCategory = pVarCategory(Index)
End Property
Public Property Let VarCategory(Index As Long, value As VariableType)
          pVarCategory(Index) = value
End Property

Public Property Get LHSType(Index As Long) As SolverInputType
          LHSType = pLHSType(Index)
End Property
Public Property Let LHSType(Index As Long, value As SolverInputType)
          pLHSType(Index) = value
End Property

Public Property Get RHSType(Index As Long) As SolverInputType
          RHSType = pRHSType(Index)
End Property
Public Property Let RHSType(Index As Long, value As SolverInputType)
          pRHSType(Index) = value
End Property

Public Property Get LHSRange(Index As Long) As Range
          Set LHSRange = pLHSRange(Index)
End Property
Public Property Set LHSRange(Index As Long, value As Range)
          Set pLHSRange(Index) = value
End Property

Public Property Get RHSRange(Index As Long) As Range
          Set RHSRange = pRHSRange(Index)
End Property
Public Property Set RHSRange(Index As Long, value As Range)
          Set pRHSRange(Index) = value
End Property

Public Property Get RHSFormula(Index As Long) As String
          RHSFormula = pRHSFormula(Index)
End Property
Public Property Let RHSFormula(Index As Long, value As String)
          pRHSFormula(Index) = value
End Property

Public Property Get RHSConstant(Index As Long) As Double
          RHSConstant = pRHSConstant(Index)
End Property
Public Property Let RHSConstant(Index As Long, value As Double)
          pRHSConstant(Index) = value
End Property

Public Property Get Relation(Index As Long) As RelationConsts
          Relation = pRelation(Index)
End Property
Public Property Let Relation(Index As Long, value As RelationConsts)
          pRelation(Index) = value
End Property

Public Property Get RowSetsBound(Index As Long) As Boolean
          RowSetsBound = pRowSetsBound(Index)
End Property
Public Property Let RowSetsBound(Index As Long, value As Boolean)
          pRowSetsBound(Index) = value
End Property

Public Property Get ConstraintSummary(Index As Long) As String
          ConstraintSummary = pConstraintSummary(Index)
End Property
Public Property Let ConstraintSummary(Index As Long, value As String)
          pConstraintSummary(Index) = value
End Property

Public Property Set Solver(NewSolver As ISolver)
          Set pSolver = NewSolver
End Property
Public Property Get Solver() As ISolver
          Set Solver = pSolver
End Property

Public Property Set SolverParameters(NewSolverParameters As Dictionary)
          Set pSolverParameters = NewSolverParameters
End Property
Public Property Get SolverParameters() As Dictionary
          Set SolverParameters = pSolverParameters
End Property

Public Property Get RHS(Index As Long) As Double
          RHS = pRHS(Index)
End Property
Public Property Let RHS(Index As Long, value As Double)
          pRHS(Index) = value
End Property

Public Property Set SparseA(Index As Long, NewSparseA As CIndexedCoeffs)
          Set pSparseA(Index) = NewSparseA
End Property
Public Property Get SparseA(Index As Long) As CIndexedCoeffs
          Set SparseA = pSparseA(Index)
End Property

Public Property Get CostCoeffs(Index As Long) As Double
790       CostCoeffs = pCostCoeffs(Index)
End Property
Public Property Let CostCoeffs(Index As Long, value As Double)
792       pCostCoeffs(Index) = value
End Property

Public Property Get RowToConstraint(Index As Long) As Long
          RowToConstraint = pRowToConstraint(Index)
End Property
Public Property Let RowToConstraint(Index As Long, value As Long)
          pRowToConstraint(Index) = value
End Property
' End of getters and setters
'===============================================================================

Function GetConstraintInstance(ByVal row As Long, ByVal constraint As Long) As Long
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

832       GetConstraintInstance = row - ConstraintToRow(constraint) + 1

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("COpenSolver", "GetConstraintInstance") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function CheckConstraintValuesAreNumeric(constraint As Long, LHSValues As Variant, RHSValues As Variant, ErrorMsg As String, AllowNonNumericIfNotConstant As Boolean) As Boolean
          ' Check that the LHS and RHS values are doubles (not vbError, or vbString, or ??)
          ' vbEmpty (blank) cell values are permitted as VBA quietly turns them into 0's in expressions
836       CheckConstraintValuesAreNumeric = False
          Dim i As Long, j As Long, instance As Long
          Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String, cell As Range
          
          ' Check LHS first; this is always an array
837       instance = 0
838       For i = 1 To UBound(LHSValues, 1)
839           For j = 1 To UBound(LHSValues, 2)
840               instance = instance + 1
                  Select Case VarType(LHSValues(i, j))
                      Case vbDouble:
                      Case vbEmpty:
                          ' This is an empty cell; VBA quietly turns this into 0 when used in an expression. We need no special handling of this, and don't report an error.
842                   Case vbError:
843                       GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
844                       ErrorMsg = "The cell '" & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "' in constraint '" & ConstraintSummary(constraint) & "' has a divide by zero error or has a problem with the function used ( eg #DIV/0! or #VALUE!)."
846                   Case Else:
                          ' It might be vbError, or vbString, or vbEmpty
847                       GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
848                       ErrorMsg = "The cell '" & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "' in constraint '" & ConstraintSummary(constraint) & "' has a value which is not numeric."
849                       On Error Resume Next
                          ' Try to show the incorrect value; just cary on with the old ErrorMsg if this gives an error
850                       ErrorMsg = "The cell '" & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "' in constraint '" & ConstraintSummary(constraint) & "' has the value '" & LHSValues(i, j) & "' which is not numeric."
                          On Error GoTo 0
852               End Select

                  If Len(ErrorMsg) > 0 Then
                      ' We have an invalid value
                      If AllowNonNumericIfNotConstant Then
                          ' Allow this if the cell depends on the variable cells - it might be valid later
                          Set cell = GetOneCellInRange(LHSCellRange, instance)
                          If Not ProperPrecedents(cell) Is Nothing Then
                              If Not Intersect(ProperPrecedents(cell), AdjustableCells) Is Nothing Then
                                  GoTo NextCellLHS
                              End If
                          End If
                      End If
                      ' Else the error is real, exit and report it
                      GoTo ExitError
                  End If
NextCellLHS:
                  ErrorMsg = vbNullString
853           Next j
854       Next i
          
855       Select Case RHSType(constraint)
              Case SolverInputType.SingleCellRange, SolverInputType.Formula, SolverInputType.constant
856               Select Case VarType(RHSValues)
                      Case vbDouble
                      Case vbEmpty:
                          ' This is an empty cell; VBA quietly turns this into 0 when used in an expression. We need no special handling of this, and don't report an error.
857                   Case vbError:
858                       GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
859                       ErrorMsg = "The right hand side in constraint '" & ConstraintSummary(constraint) & "' has a divide by zero error or has a problem with the function used (eg #DIV/0! or #VALUE!)."
860                   Case Else:
                          ' It might be vbError, or vbString, or anything else?
862                       GetConstraintInstanceData constraint, 1, LHSCellRange, RHSCellRange, RHSstring
863                       ErrorMsg = "The right hand side in constraint '" & ConstraintSummary(constraint) & "' has a value which is not a single numeric value."
                          ' Try to show the incorrect value; just cary on with the old ErrorMsg if this gives an error
864                       On Error Resume Next
865                       ErrorMsg = "The right hand side in constraint '" & ConstraintSummary(constraint) & "' has the value '" & RHSValues & "' which is not numeric."
                          On Error GoTo 0
867               End Select

                  If Len(ErrorMsg) > 0 Then
                      ' We have an invalid value
                      If AllowNonNumericIfNotConstant Then
                          ' Allow this if the cell depends on the variable cells - it might be valid later
                          Set cell = RHSCellRange
                          If Not ProperPrecedents(cell) Is Nothing Then
                              If Not Intersect(ProperPrecedents(cell), AdjustableCells) Is Nothing Then
                                  GoTo NextCellSingleRHS
                              End If
                          End If
                      End If
                      ' Else the error is real, exit and report it
                      GoTo ExitError
                  End If
NextCellSingleRHS:
868           Case SolverInputType.MultiCellRange
869               instance = 0
870               For i = 1 To UBound(RHSValues, 1)
871                   For j = 1 To UBound(RHSValues, 2)
872                       instance = instance + 1
873                       Select Case VarType(RHSValues(i, j))
                              Case vbDouble:
                              Case vbEmpty:
                                  ' This is an empty cell; VBA quietly turns this into 0 when used in an expression. We need no special handling of this, and don't report an error.
874                           Case vbError:
875                               GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
876                               ErrorMsg = "The right hand side in constraint '" & ConstraintSummary(constraint) & "' has a divide by zero error or has a problem with the function used ( eg #DIV/0! or #VALUE!)."
877                               GoTo ExitError
878                           Case Else:
                                  ' It might be vbError, or vbString, or vbEmpty
879                               GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
880                               ErrorMsg = "The cell '" & RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "' in constraint '" & ConstraintSummary(constraint) & "' has a value which is not numeric."
881                               On Error Resume Next
882                               ErrorMsg = "The cell '" & RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "' in constraint '" & ConstraintSummary(constraint) & "' has the value '" & RHSValues(i, j) & "' which is not numeric."
                                  On Error GoTo 0
884                       End Select

                          If Len(ErrorMsg) > 0 Then
                              ' We have an invalid value
                              If AllowNonNumericIfNotConstant Then
                                  ' Allow this if the cell depends on the variable cells - it might be valid later
                                  Set cell = GetOneCellInRange(RHSCellRange, instance)
                                  If Not ProperPrecedents(cell) Is Nothing Then
                                      If Not Intersect(ProperPrecedents(cell), AdjustableCells) Is Nothing Then
                                          GoTo NextCellMultiRHS
                                      End If
                                  End If
                              End If
                              ' Else the error is real, exit and report it
                              GoTo ExitError
                          End If
NextCellMultiRHS:
                          ErrorMsg = vbNullString
885                   Next j
886               Next i
887       End Select
888       CheckConstraintValuesAreNumeric = True
          Exit Function
          
ExitError:
          ErrorMsg = ErrorMsg & vbNewLine & vbNewLine & _
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
          RaiseError = False
          On Error GoTo ErrorHandler

889       If LHSType(constraint) = SolverInputType.SingleCellRange Then
              ' Create a 1x1 array
890           LHSValues = Create1x1Array(LHSRange(constraint).Value2)
891       Else ' LHSType(constraint) = SolverInputType.MultiCellRange
              ' Copy all the values into a variant, which will then be an
              ' array of variants; see http://support.microsoft.com/kb/213798
892           LHSValues = LHSRange(constraint).Value2
893       End If

894       If RHSType(constraint) = SolverInputType.SingleCellRange Then
              ' A variant containing just a single double
895           RHSValues = RHSRange(constraint).Value2
896       ElseIf RHSType(constraint) = SolverInputType.MultiCellRange Then
              ' Copy all the values into a variant, which will then be an
              ' array of variants
897           RHSValues = RHSRange(constraint).Value2
898       ElseIf RHSType(constraint) = SolverInputType.Formula Then
              ' Will not throw an error, but may put an error value in RHSValues
899           RHSValues = sheet.Evaluate("=" & RHSFormula(constraint))
900       Else  ' RHSType(constraint) = SolverInputType.Constant
901           RHSValues = RHSConstant(constraint)
902       End If

1171      If ValidateNumeric Then
              Dim ErrorMsg As String
              If Not CheckConstraintValuesAreNumeric(constraint, LHSValues, RHSValues, ErrorMsg, _
                                                     AllowNonNumericIfNotConstant) Then
1173              Err.Raise Number:=OpenSolver_BuildError, Description:=ErrorMsg
              End If
1174      End If

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("COpenSolver", "GetCurrentConstraintValues") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Function GetCurrentObjectiveValue(Optional ValidateNumeric As Boolean = True, Optional AllowNonNumericIfNotConstant As Boolean = False) As Variant
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          If ObjRange Is Nothing Then
              GetCurrentObjectiveValue = 0
              GoTo ExitFunction
          End If
              
          Dim ObjValue As Variant
          ObjValue = ObjRange.Value2
          
          If ValidateNumeric Then
              Dim ErrorMsg As String
              ' Check if objective has a value that is not a number
              If VarType(ObjValue) <> vbDouble Then
                  If VarType(ObjValue) = vbError Then
                      ErrorMsg = "The objective cell appears to contain an error. This could have occurred if there is a divide by zero error or if you have used the wrong function (eg #DIV/0! or #VALUE!)."
                  Else
                      ErrorMsg = "The objective cell does not appear to contain a numeric value."
                  End If
              End If
              
              If Len(ErrorMsg) > 0 Then
                  If AllowNonNumericIfNotConstant Then
                      If Not ProperPrecedents(ObjRange) Is Nothing Then
                          If Not Intersect(ProperPrecedents(ObjRange), AdjustableCells) Is Nothing Then
                              GoTo DontRaise
                          End If
                      End If
                  End If
                  ErrorMsg = _
                      ErrorMsg & vbNewLine & vbNewLine & _
                      IIf(AllowNonNumericIfNotConstant, "This can happen if you have made an error when constructing the model.", _
                                                        "This can happen if your model is nonlinear, or if you have made an error when constructing the model.") & _
                      " Please fix this and try again."
                  Err.Raise OpenSolver_BuildError, Description:=ErrorMsg
              End If
          End If
DontRaise:
          GetCurrentObjectiveValue = ObjValue

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("COpenSolver", "GetCurrentObjectiveValue") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function


' Get the information required to display a single instance of a constraint to the user
Sub GetConstraintInstanceData(constraint As Long, instance As Long, LHSCellRange As Range, RHSCellRange As Range, RHSstring As String)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

903       Select Case LHSType(constraint)
              Case SolverInputType.SingleCellRange
904               Set LHSCellRange = LHSRange(constraint)
905           Case SolverInputType.MultiCellRange
906               Set LHSCellRange = GetOneCellInRange(LHSRange(constraint), instance)
907           Case SolverInputType.Formula
908               Debug.Assert False ' should never happen
909           Case SolverInputType.constant
910               Debug.Assert False ' should never happen
911       End Select
912       Select Case RHSType(constraint)
              Case SolverInputType.SingleCellRange
913               Set RHSCellRange = RHSRange(constraint)
914               RHSstring = ""
915           Case SolverInputType.MultiCellRange
916               Set RHSCellRange = GetOneCellInRange(RHSRange(constraint), instance)
917               RHSstring = ""
918           Case SolverInputType.Formula, SolverInputType.constant
919               Set RHSCellRange = Nothing
920               RHSstring = RHSFormula(constraint)
924       End Select

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("COpenSolver", "GetConstraintInstanceData") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Sub BuildModelFromSolverData(Optional LinearityOffset As Double = 0, _
                             Optional ShouldCheckLinearity As Boolean = False, _
                             Optional ShouldMinimiseUserInteraction As Boolean = False, _
                             Optional ShouldSolveRelaxation As Boolean = False, _
                             Optional modelSheet As Worksheet)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

935       SolveStatus = OpenSolverResult.Unsolved
936       SolveStatusString = "Unsolved"
937       SolveStatusComment = ""
940       SolutionWasLoaded = False
941       ModelStatus = Unitialized

          MinimiseUserInteraction = ShouldMinimiseUserInteraction

          ' Change to manual calculation mode
          Dim oldCalculationMode As Long
969       oldCalculationMode = Application.Calculation
970       Application.Calculation = xlCalculationManual

          Dim ScreenStatus As Boolean
          ScreenStatus = Application.ScreenUpdating
971       Application.ScreenUpdating = False

972       Application.Cursor = xlWait

          Dim StartTimeOfBuild As Single, StartDateTimeOfBuild As Date
          Dim EndTimeOfBuild As Single, EndDateTimeOfBuild As Date
944       StartTimeOfBuild = Timer    ' Warning: This goes back to zero at midnight!
945       StartDateTimeOfBuild = Now()

947       Application.EnableCancelKey = xlErrorHandler
          
          GetActiveSheetIfMissing modelSheet
          Set sheet = modelSheet
          
          ' Find the solver
960       Set Solver = CreateSolver(GetChosenSolver(sheet))
          LastUsedSolver = Solver.ShortName

          '--------------------------------------------------------------------------
          ' ADJUSTABLE CELLS SETUP
          
          ' Remove any overlap
980       Set AdjustableCells = GetDecisionVariablesNoOverlap(sheet)
          ' Handle any merged ranges
          Dim BadCell As Range
981       If Not CheckRangeContainsNoAmbiguousMergedCells(AdjustableCells, BadCell) Then
983           Err.Raise Number:=OpenSolver_BuildError, Description:="This model contains a decision variable " & BadCell.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & " that is inaccessible as it is within the merged range " & BadCell.MergeArea.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & ". Please redefine the decision variable cells, and try again."
984           GoTo ExitSub
985       End If
          ' Check we have any cells after all this
986       NumVars = AdjustableCells.Count
987       If NumVars < 1 Then
989           Err.Raise Number:=OpenSolver_BuildError, Description:="There are no decision variable cells in this model, and so the model cannot be solved."
990       End If

          ' Get names for all the variables, creating names that distinguish cells on different sheets
991       ReDim pVarName(1 To NumVars) As String
          ReDim pVarCategory(1 To NumVars) As VariableType
          Set VarNameToIndex = New Collection
          
          Dim AdjCell As Range, i As Long
996       i = 1
997       For Each AdjCell In AdjustableCells
999           VarName(i) = GetCellName(AdjCell)
              VarCategory(i) = VariableType.VarContinuous
              VarNameToIndex.Add Item:=i, Key:=VarName(i)
              i = i + 1
1003      Next AdjCell
          
          ' Reset the cell values to catch protected sheets
1004      TestCellsForWriting AdjustableCells
          
          ' Initial sheet calculation
          If Not ForceCalculate("Warning: The initial worksheet calculation did not complete, and so the model may not be generated correctly. Would you like to try again?", MinimiseUserInteraction) Then
1008          On Error GoTo ErrorHandler
1009          Err.Raise Number:=OpenSolver_BuildError, Description:="Calculation failed. The model has not been built so cannot be solved."
1010      End If

          '--------------------------------------------------------------------------
          ' OBJECTIVE SETUP
1011      ObjectiveSense = GetObjectiveSense(sheet)
1013      ObjectiveTargetValue = GetObjectiveTargetValue(sheet)
          Set ObjRange = GetObjectiveFunctionCell(sheet, Validate:=True)
          ' Check that objective is not an non-numeric constant
          GetCurrentObjectiveValue AllowNonNumericIfNotConstant:=True

          '--------------------------------------------------------------------------
          ' MODEL OPTIONS SETUP
1041      AssumeNonNegativeVars = GetNonNegativity(sheet)
          ShowIterationResults = GetShowSolverProgress(sheet)
          
          ' Collate all solver parameters (ours and user-defined)
          Set SolverParameters = GetSolverParametersDict(Solver, sheet)

          ' Check for Solver's "Ignore Integer Constraints" option unless solving relaxation
          If Not ShouldSolveRelaxation Then
              ' If we are in Excel 2010 or newer, and the user has set "Ignore Integer Constraints", then we report an error
1266          If Val(Application.Version) >= 14 And GetIgnoreIntegerConstraints(sheet) Then
                   If MinimiseUserInteraction Then GoTo RaiseError
                   If Not MsgBox("You have the 'Ignore integer constraints' option enabled in Solver, so OpenSolver cannot proceed with solving. You can either disable this option, or use the 'Solve Relaxation' option in the OpenSolver menu. Would you like OpenSolver to disable this option and proceed?", vbCritical + vbYesNo + vbDefaultButton1, "OpenSolver Warning") = vbNo Then
                       SetIgnoreIntegerConstraints False, sheet
                   Else
RaiseError:
                       ' Set UserCancelledError because the user has already confirmed the abort if not in silent mode
                       Err.Raise Number:=OpenSolver_UserCancelledError, Description:="You have the Solver option 'ignore integer constraints' turned on. Please turn this off in the Solver (not OpenSolver) options menu and then solve with OpenSolver again or choose OpenSolver's solve relaxed model option"
                   End If
1271          End If
          End If
          
          '--------------------------------------------------------------------------
          ' CONSTRAINTS SETUP
1048      NumConstraints = GetNumConstraints(sheet)
1049      UpdateStatusBar "OpenSolver: Setting Up Problem... " & NumVars & " vars, " & NumConstraints & " Solver constraints", True
          'Set up arrays to be the correct size. We allow no constraints!
1050      If NumConstraints > 0 Then
1054          ReDim pLHSType(NumConstraints) As SolverInputType
1055          ReDim pLHSRange(NumConstraints) As Range
              ReDim pRelation(NumConstraints) As RelationConsts
1053          ReDim pRHSType(NumConstraints) As SolverInputType
1056          ReDim pRHSRange(NumConstraints) As Range
1057          ReDim pRHSFormula(NumConstraints) As String
              ReDim pRHSConstant(NumConstraints) As Double
1058          ReDim pConstraintSummary(NumConstraints) As String
1059          ReDim ConstraintToRow(NumConstraints) As Long
              ReDim RowCount(NumConstraints) As Long
1060      End If
1061      Set VarLowerBounds = New Collection
1062      Set VarUpperBounds = New Collection
1063      ReDim pRowSetsBound(1)

1064      NumRows = 0
          Dim constraint As Long, row As Long
1065      row = 1
          ' Iterate through each solver constraint
1066      For constraint = 1 To NumConstraints
1067          UpdateStatusBar "OpenSolver: Setting Up Constraints... " & NumVars & " vars, " & constraint & "/" & NumConstraints & " Solver constraints"
                           
              ' LEFT HAND SIDE
              Dim rangeLHS As Range, sRefersToLHS As String
              Set rangeLHS = GetConstraintLhs(constraint, sheet, Validate:=True, RefersTo:=sRefersToLHS)
              Dim LHSCount As Long
1082          LHSCount = rangeLHS.Count
              
              ' RELATIONSHIP
              Dim rel As Long
              rel = GetConstraintRel(constraint, sheet)
              Relation(constraint) = rel
              
1083          Select Case rel
              Case RelationINT, RelationBIN
                  ' Check that they haven't set as integer or binary something that is not a decision variable...
                  Dim intersection As Range
1084              Set intersection = Intersect(rangeLHS, AdjustableCells)
1085              If intersection Is Nothing Then
                      ' The cells are not all decision variables
1086                  On Error GoTo ErrorHandler
1087                  Err.Raise Number:=OpenSolver_BuildError, Description:="A range of cells are specified as bin or int that are not decision variables. OpenSolver does not support this yet."
1088              End If
1089              If intersection.Count = rangeLHS.Count Then
1090                  If rel = RelationINT Then
1091                      Set IntegerCellsRange = ProperUnion(IntegerCellsRange, rangeLHS)
1092                  Else
1093                      Set BinaryCellsRange = ProperUnion(BinaryCellsRange, rangeLHS)
1094                  End If
1095              Else
                      ' The cells are not all decision variables
1096                  On Error GoTo ErrorHandler
1097                  Err.Raise Number:=OpenSolver_BuildError, Description:="Unable to find a cell specified as bin or int in the decision variables."
1098              End If
                  ' This constraint does not relate to any new rows, so mark it
                  ' with a dummy entry. row = same value as we get for the
                  ' next entry, so this refers to no rows.
1099              ConstraintToRow(constraint) = row
              
              ' RIGHT HAND SIDE
              Case RelationGE, RelationLE, RelationEQ
                  ' The constraint is a full equation with a RHS
                  Dim valRHS As Double, rangeRHS As Range, sRefersToRHS As String, RHSRefersToFormula As Boolean
1101              Set rangeRHS = GetConstraintRhs(constraint, sRefersToRHS, valRHS, RHSRefersToFormula, sheet)
                  
1115              ConstraintSummary(constraint) = StripWorksheetNameAndDollars(sRefersToLHS, sheet) & _
                                                  " " & RelationEnumToString(rel) & " " & _
                                                  StripWorksheetNameAndDollars(sRefersToRHS, sheet)
                                                  
                  Dim RHSCount As Long
1121              If Not rangeRHS Is Nothing Then
                      RHSCount = rangeRHS.Count
                  Else
                      RHSCount = 1
                  End If
                  
                  ' Check we have a compatible constraint system.
1122              If (LHSCount <> 1 And RHSCount <> 1) And (LHSCount <> RHSCount) Then
1123                  On Error GoTo ErrorHandler
1124                  Err.Raise Number:=OpenSolver_BuildError, Description:="The constraint '" & ConstraintSummary(constraint) & "' has a different cell count on the left and the right. The model cannot be built."
1125              End If
                  
                  ' Store left hand side:
1126              Set LHSRange(constraint) = rangeLHS
1127              LHSType(constraint) = IIf(LHSCount = 1, SolverInputType.SingleCellRange, SolverInputType.MultiCellRange)
                  
                  ' Store right hand side:
1134              If Not rangeRHS Is Nothing Then
1135                  Set RHSRange(constraint) = rangeRHS
1136                  RHSType(constraint) = IIf(RHSCount = 1, SolverInputType.SingleCellRange, SolverInputType.MultiCellRange)
1138              ElseIf RHSRefersToFormula Then
                      ' Store the formula internally as a string
1139                  RHSFormula(constraint) = sRefersToRHS
1140                  RHSType(constraint) = SolverInputType.Formula
1142              Else
                      ' Store a constant
                      RHSFormula(constraint) = sRefersToRHS
1143                  RHSType(constraint) = SolverInputType.constant
                      RHSConstant(constraint) = valRHS
1145              End If

                  ' Check that all constant values are numeric
                  ' We allow non-numeric values that might later change with the variable cells
                  Dim LHSValues As Variant
                  Dim RHSValues As Variant
                  GetCurrentConstraintValues constraint, LHSValues, RHSValues, AllowNonNumericIfNotConstant:=True
                  
                  ' Check if this constraint applies any bounds on decision variables
1146              If LHSCount = 1 Then
                      ' Single LHS, single RHS
1147                  ApplyBounds row, rel, rangeLHS, rangeRHS, valRHS, RHSRefersToFormula
1148              Else
                      Dim CurRow As Long
1154                  If RHSCount > 1 Then
                          ' multiple LHS, multiple RHS
1155                      For CurRow = 1 To LHSCount
1156                          ApplyBounds row + CurRow - 1, rel, rangeLHS(CurRow), rangeRHS(CurRow), valRHS, RHSRefersToFormula
1157                      Next CurRow
1158                  Else
                          ' multiple LHS, single RHS
1159                      For CurRow = 1 To LHSCount
1160                          ApplyBounds row + CurRow - 1, rel, rangeLHS(CurRow), rangeRHS, valRHS, RHSRefersToFormula
1161                      Next CurRow
1162                  End If
1163              End If
                  
                  ' Store the row/constraint information
                  RowCount(constraint) = LHSCount
1169              ConstraintToRow(constraint) = row
1168              row = row + LHSCount
              
              Case Else
                  ' Don't allow 'alldiff' from Excel 2010 or any other wierd things
1116              On Error GoTo ErrorHandler
1118              Err.Raise Number:=OpenSolver_BuildError, Description:="Unrecognised relationship for constraint " & ConstraintSummary(constraint)
1175          End Select
1176      Next constraint
          NumRows = row - 1
          
          If NumRows > 0 Then
              ' Make the row to constraint mapping
              ReDim pRowToConstraint(NumRows) As Long
              row = 0
              For constraint = 1 To NumConstraints
                  For i = 1 To RowCount(constraint)
                      RowToConstraint(row + i) = constraint
                  Next i
                  row = row + RowCount(constraint)
              Next constraint
          End If

          ' Binary trumps integer, so remove all overlap
          Set IntegerCellsRange = SetDifference(IntegerCellsRange, BinaryCellsRange)
          
          If Not IntegerCellsRange Is Nothing Then
              Me.NumIntVars = IntegerCellsRange.Count
              For Each AdjCell In IntegerCellsRange
                  Me.VarCategory(Me.VarNameToIndex(GetCellName(AdjCell))) = VarInteger
              Next AdjCell
          End If
          If Not BinaryCellsRange Is Nothing Then
              Me.NumBinVars = BinaryCellsRange.Count
              For Each AdjCell In BinaryCellsRange
                  Me.VarCategory(Me.VarNameToIndex(GetCellName(AdjCell))) = VarBinary
              Next AdjCell
          End If
          Me.NumDiscreteVars = Me.NumIntVars + Me.NumBinVars

1177      If NumRows > 0 Then
1178          ReDim Preserve pRowSetsBound(NumRows)
1179      End If
          
          'Set up extra model info depending on the solver
1180      Select Case Solver.ModelType
          Case Diff
1181          If NumRows > 0 Then
                  ReDim Preserve pSparseA(NumRows) As CIndexedCoeffs
                  ReDim pRHS(NumRows) As Double
              End If
              ReDim pCostCoeffs(NumVars) As Double
              
              Set DiffModel = New CModelDiff
              If Not DiffModel.ProcessSolverModel(Me, LinearityOffset, ShouldCheckLinearity) Then GoTo ExitSub
          Case Parsed
              ' Parse spreadsheet and build model
              Set ParsedModel = New CModelParsed
              If Not ParsedModel.ProcessSolverModel(Me) Then GoTo ExitSub
1188      End Select

1189      ModelStatus = Built
1190      EndTimeOfBuild = Timer
1191      EndDateTimeOfBuild = Now()
1192      SecondsTakenToBuildModel = DateDiff("s", StartDateTimeOfBuild, EndDateTimeOfBuild)
1193      TimeTakenToBuildModel = EndTimeOfBuild - StartTimeOfBuild

ExitSub:
          Application.Cursor = xlDefault
          Application.StatusBar = False
          Application.ScreenUpdating = ScreenStatus
          Application.Calculation = oldCalculationMode
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("COpenSolver", "BuildModelFromSolverData") Then Resume
          RaiseError = True
          GoTo ExitSub
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
          RaiseError = False
          On Error GoTo ErrorHandler

1217      If rel <> RelationGE And rel <> RelationLE Then GoTo ExitSub
          
          ' Check if we need upper bounds
1220      If Not SolverUsesUpperBounds(Solver.ShortName) And rel = RelationLE Then GoTo ExitSub
           
          ' Following Solver, we don't apply formulae as lower bounds, even if constant
1223      If RHSRefersToFormula Then GoTo ExitSub
          
          ' Check if LHS cell is a variable cell
1226      If Intersect(LHS, AdjustableCells) Is Nothing Then GoTo ExitSub
          
1229      If Not RHS Is Nothing Then
              If IsNumeric(RHS.value) Then
1230              If Not RHS.HasFormula Then
                      ' The RHS range is a constant - apply a bound
1231                  UpdateBound row, GetCellName(LHS), CDbl(RHS.value), rel
1232              Else
                      ' The RHS range is a formula, check if this depends on the decision vars
                      ' Will raise an error is no precedents are found (e.g. if they are on another sheet
1234                  If Intersect(RHS.Precedents, AdjustableCells) Is Nothing Then
1235                      UpdateBound row, GetCellName(LHS), CDbl(RHS.value), rel
                      End If
1236              End If
1237          End If
1238      Else
              ' The RHS is a numeric constant - apply a bound
1239          UpdateBound row, GetCellName(LHS), RHSValue, rel
1240      End If
          
ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          ' Ignore a 'No cells were found' precedents error
          If Err.Number = 1004 Then
              GoTo ExitSub
          End If

          If Not ReportError("COpenSolver", "ApplyBounds") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Sub UpdateBound(row As Long, VarName As String, bound As Double, rel As Long)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim SetBound As Boolean
          SetBound = True
          
          ' Only GE constraints can provide lower bounds
1246      If rel = RelationGE Then
              ' Remove existing lower bound on variable if our new one is better
1247          If TestKeyExists(VarLowerBounds, VarName) Then
1248              If bound > VarLowerBounds(VarName) Then
1249                  VarLowerBounds.Remove VarName
1250              Else
                      SetBound = False
                  End If
1251          End If
              
1252          If SetBound Then VarLowerBounds.Add bound, VarName
          ' Only LE constraints can provide upper bounds
1253      ElseIf rel = RelationLE Then
              ' Remove existing upper bound on variable if our new one is better
1254          If TestKeyExists(VarUpperBounds, VarName) Then
1255              If bound < VarUpperBounds(VarName) Then
1256                  VarUpperBounds.Remove VarName
                  Else
                      SetBound = False
1257              End If
1258          End If
              
1259          If SetBound Then VarUpperBounds.Add bound, VarName
1260      End If
          
          ' Track that this constraint row applies a bound
1261      If UBound(pRowSetsBound) < row Then ReDim Preserve pRowSetsBound(row)
1262      RowSetsBound(row) = True

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("COpenSolver", "UpdateBound") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Sub ReportAnySolutionSubOptimality()
          Select Case SolveStatus
          Case OpenSolverResult.Optimal                ' We are optimal
          Case OpenSolverResult.Pending                ' The solve is scheduled
          Case OpenSolverResult.NotLinear              ' Already shown non-linearity details
          Case OpenSolverResult.AbortedThruUserAction  ' Already shown escape dialog
       
          Case Else
              MsgBox "OpenSolver could not find an optimal solution, and reported:" & vbCrLf & _
                     SolveStatusString & vbCrLf & vbCrLf & _
                     IIf(SolveStatusComment = "", "", SolveStatusComment & vbCrLf & vbCrLf) & _
                     IIf(SolutionWasLoaded, "The solution generated has been loaded into the spreadsheet.", "No solution was available to load into the spreadsheet.")
          End Select
End Sub

Sub ClearQuickSolve()
1992      Set QuickSolve = Nothing
End Sub

Function InitializeQuickSolve(ParamRange As Range) As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

1993      InitializeQuickSolve = False
1994      If ModelStatus <> Built Then
1995          Err.Raise OpenSolver_BuildError, Description:="QuickSolve Initialisation Failed: The model cannot be solved as it has not yet been built."
1997      End If

1998      Set QuickSolve = New CQuickSolve
1999      If QuickSolve.AnalyseParameters(Me, ParamRange) Then
2000          InitializeQuickSolve = True
2001      Else
2002          Set QuickSolve = Nothing
2003      End If

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("COpenSolver", "InitializeQuickSolve") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Sub DoQuickSolve(Optional SolveRelaxation As Boolean = False, Optional MinimiseUserInteraction As Boolean = False)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
2013      If QuickSolve Is Nothing Then
2014          Err.Raise OpenSolver_ModelError, Description:="Quick Solve has not been initialized. Please choose Intialize Quick Solve, and then try again."
2015      Else
2016          If QuickSolve.UpdateModelRHS(sheet, pRHS) Then
2017              SolveModel Me, SolveRelaxation, MinimiseUserInteraction
2018          End If
2019      End If

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("COpenSolver", "DoQuickSolve") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Public Sub PrepareForSolution()
          ReDim pVarFinalValue(NumVars) As Double
1435      ReDim pVarCellName(NumVars) As String
1436      If bGetDuals Then
1437          ReDim pConFinalValue(NumRows) As Double
1438          ReDim pConShadowPrice(NumRows) As Double
1439          ReDim pConIncrease(NumRows) As Double
1440          ReDim pConDecrease(NumRows) As Double
1441          ReDim pVarReducedCost(NumVars) As Double
1442          ReDim pVarIncrease(NumVars) As Double
1443          ReDim pVarDecrease(NumVars) As Double
1444      End If
End Sub

Public Sub LoadResultsToSheet()
   If SolutionWasLoaded Then
        AdjustableCells.Value2 = 0
        Dim i As Long
        For i = 1 To NumVars
            If Len(VarCellName(i)) <> 0 Then
                AdjustableCells.Worksheet.Range(VarCellName(i)).Value2 = VarFinalValue(i)
            End If
        Next i
    End If
End Sub
