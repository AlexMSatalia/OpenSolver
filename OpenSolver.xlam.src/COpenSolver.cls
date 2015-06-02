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

Public sheetName As String ' Sheetname in quotes as used for prefixing
Dim RawSheetName As String ' Just the worksheet name
Dim RawWorkBookName As String ' Just the workbook name
Public book As Workbook
Public sheet As Worksheet

' Whether to show user alerts
Public MinimiseUserInteraction As Boolean
Public SolveRelaxation As Boolean

Dim pSolverParameters As Dictionary
Public ShowIterationResults As Boolean

' Data for our model. We work thru the original Excel 'constraints' and form model 'rows' from these
Public AdjustableCells As Range    ' A range we form containing all the adjustable (ie decision) variables
Public BinaryCellsRange As Range   ' A range we form containing all the adjustable (ie decision) variables that are binary
Public IntegerCellsRange As Range  ' A range we form containing all the adjustable (ie decision) variables that are integer (which may include binary variables)
Public AssumeNonNegativeVars As Boolean   ' Excel stores ...!solver_neg=1 if Assume Linear Model is turned on, 2 if off (NB: Not 0!)

' For the LHS and RHS of each Excel constraint, we store its type, and its LHS and RHS
Dim pLHSType() As SolverInputType
Dim pLHSRange() As Range
Dim pRELs() As RelationConsts
Dim pRHSType() As SolverInputType
Dim pRHSRange() As Range
Dim pRHSFormula() As String
Dim pConstraintSummary() As String

' The original values of the LHS and RHS of the Excel constraints before we start perturbing the adjustable cells;
' each entry is stored as a variant that could be a double or an array
Dim pLHSOriginalValues() As Variant
Dim pRHSOriginalValues() As Variant

Public ObjRange As Range

' Details for the model that we build.
Public numVars As Long, NumIntVars As Long, NumBinVars As Long, NumConstraints As Long
Public NumRows As Long
Public objValue As Double  ' The constant value of the objective
Public ObjectiveSense As Long
Public ObjectiveTargetValue As Double  ' Used when the user wants the objective to achieve some stated target value
Dim pVarNames() As String
Dim pCostCoeffs() As Double
Public ObjectiveFunctionConstant As Double
Dim pRHS() As Double ' The RHS values for each row in the LP
Dim pRelation() As RelationConsts ' For each row, we remember its relation (=, <=, >=)
Public VarLowerBounds As Collection
Public VarUpperBounds As Collection

'Details for the chosen solver
Dim pSolver As ISolver
Public ExternalSolverPathName As String 'Path to solver
Public LogFilePathName As String

' Details we need for non-linear solver
Dim pRowSetsBound() As Boolean

' Details for sensitivity analysis
Public bGetDuals As Boolean ' Whether to get duals
Public rConstraintList As Range
Public DualsOnNewSheet As Boolean
Public DualsOnSameSheet As Boolean

' Don't use this directly, use the property methods
Dim pFinalValue() As Double
Dim pShadowPrice() As Double
Dim pIncreaseCon() As Double
Dim pDecreaseCon() As Double
Dim pFinalVarValue() As Double
Dim pReducedCosts() As Double
Dim pIncreaseVar() As Double
Dim pDecreaseVar() As Double
Dim pVarCell() As String

' This is the A matrix for our Solver model
Dim pSparseA() As CIndexedCoeffs

Dim pParsedModel As CModelParsed

' We remember how to turn a row index in the LP back into a constraint index in the Excel list of constraints
' This array stores the first row in the LP associated with each Excel constraint
Dim MappingRowsToConstraints() As Long

' Our result as an enum and a string
Public SolveStatus As OpenSolverResult
Public SolveStatusString As String    ' Eg Infeasible, Opimal etc
Public SolveStatusComment As String ' Eg constraint 1 cannot be satisified
Public SolutionWasLoaded As Boolean ' True if we loaded a solution into the spreadsheet

Public ModelStatus As OpenSolver_ModelStatus

' The following are calculated differently; TimeTakenToBuildModel fails if we cross midnight!
Dim TimeTakenToBuildModel As Single, SecondsTakenToBuildModel As Long

' Support for doing a quick-solve where we only modify the RHS
Dim QuickSolve As CQuickSolve

' ============================================================================
' Array accessors
Public Property Get CostCoeffs(Index As Long) As Double
790       CostCoeffs = pCostCoeffs(Index)
End Property

Public Property Let CostCoeffs(Index As Long, value As Double)
791       If Index > UBound(pCostCoeffs) Then ReDim Preserve pCostCoeffs(Index)
792       pCostCoeffs(Index) = value
End Property

Public Property Get FinalVarValue(Index As Long) As Double
793       FinalVarValue = pFinalVarValue(Index)
End Property

Public Property Let FinalVarValue(Index As Long, value As Double)
794       If Index > UBound(pFinalVarValue) Then ReDim Preserve pFinalVarValue(Index)
795       pFinalVarValue(Index) = value
End Property

Public Property Get FinalValue(Index As Long) As Double
796       FinalValue = pFinalValue(Index)
End Property

Public Property Let FinalValue(Index As Long, value As Double)
797       If Index > UBound(pFinalValue) Then ReDim Preserve pFinalValue(Index)
798       pFinalValue(Index) = value
End Property

Public Property Get ShadowPrice(Index As Long) As Double
799       ShadowPrice = pShadowPrice(Index)
End Property

Public Property Let ShadowPrice(Index As Long, value As Double)
800       If Index > UBound(pShadowPrice) Then ReDim Preserve pShadowPrice(Index)
801       pShadowPrice(Index) = value
End Property

Public Property Get ReducedCosts(Index As Long) As Double
802       ReducedCosts = pReducedCosts(Index)
End Property

Public Property Let ReducedCosts(Index As Long, value As Double)
803       If Index > UBound(pReducedCosts) Then ReDim Preserve pReducedCosts(Index)
804       pReducedCosts(Index) = value
End Property

Public Property Get VarCell(Index As Long) As String
805       VarCell = pVarCell(Index)
End Property

Public Property Let VarCell(Index As Long, value As String)
806       If Index > UBound(pVarCell) Then ReDim Preserve pVarCell(Index)
807       pVarCell(Index) = value
End Property

Public Property Get IncreaseCon(Index As Long) As Double
808       IncreaseCon = pIncreaseCon(Index)
End Property

Public Property Let IncreaseCon(Index As Long, value As Double)
809       If Index > UBound(pIncreaseCon) Then ReDim Preserve pIncreaseCon(Index)
810       pIncreaseCon(Index) = value
End Property

Public Property Get DecreaseCon(Index As Long) As Double
811       DecreaseCon = pDecreaseCon(Index)
End Property

Public Property Let DecreaseCon(Index As Long, value As Double)
812       If Index > UBound(pDecreaseCon) Then ReDim Preserve pDecreaseCon(Index)
813       pDecreaseCon(Index) = value
End Property

Public Property Get IncreaseVar(Index As Long) As Double
814       IncreaseVar = pIncreaseVar(Index)
End Property

Public Property Let IncreaseVar(Index As Long, value As Double)
815       If Index > UBound(pIncreaseVar) Then ReDim Preserve pIncreaseVar(Index)
816       pIncreaseVar(Index) = value
End Property

Public Property Get DecreaseVar(Index As Long) As Double
817       DecreaseVar = pDecreaseVar(Index)
End Property

Public Property Let DecreaseVar(Index As Long, value As Double)
818       If Index > UBound(pDecreaseVar) Then ReDim Preserve pDecreaseVar(Index)
819       pDecreaseVar(Index) = value
End Property

Public Property Get VarNames(Index As Long) As String
820       VarNames = pVarNames(Index)
End Property

Public Property Let VarNames(Index As Long, value As String)
          pVarNames(Index) = value
End Property

Public Property Get LHSOriginalValues(Index As Long) As Variant
          LHSOriginalValues = pLHSOriginalValues(Index)
End Property

Public Property Let LHSOriginalValues(Index As Long, value As Variant)
          pLHSOriginalValues(Index) = value
End Property

Public Property Get RHSOriginalValues(Index As Long) As Variant
          RHSOriginalValues = pRHSOriginalValues(Index)
End Property

Public Property Let RHSOriginalValues(Index As Long, value As Variant)
          pRHSOriginalValues(Index) = value
End Property

Public Property Get RHS(Index As Long) As Double
          RHS = pRHS(Index)
End Property

Public Property Let RHS(Index As Long, value As Double)
          pRHS(Index) = value
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

Public Property Get RELs(Index As Long) As RelationConsts
          RELs = pRELs(Index)
End Property

Public Property Let RELs(Index As Long, value As RelationConsts)
          pRELs(Index) = value
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

Public Property Set ParsedModel(NewParsedModel As CModelParsed)
          Set pParsedModel = NewParsedModel
End Property

Public Property Get ParsedModel() As CModelParsed
          Set ParsedModel = pParsedModel
End Property

Public Property Set SolverParameters(NewSolverParameters As Dictionary)
          Set pSolverParameters = NewSolverParameters
End Property

Public Property Get SolverParameters() As Dictionary
          Set SolverParameters = pSolverParameters
End Property

Public Property Set SparseA(Index As Long, NewSparseA As CIndexedCoeffs)
          Set pSparseA(Index) = NewSparseA
End Property

Public Property Get SparseA(Index As Long) As CIndexedCoeffs
          Set SparseA = pSparseA(Index)
End Property
' End of getters and setters
'===============================================================================


Public Function GetNumRows() As Long
824       GetNumRows = NumRows
End Function

Public Function GetSparseACount(Index As Long) As Long
825       GetSparseACount = SparseA(Index).Count
End Function

Sub GetConstraintFromRow(row As Long, ByRef constraint As Long, ByRef instance As Long)
          ' Get the Excel constraint that goes with this LP row
          ' The incoming Constraint value will be either 1 or the constraint found on the last call to this
          ' We return the entry where MappingRowsToConstraints(Constraint) <= Row < MappingRowsToConstraints(Constraint+1)
          ' Note that MappingRowsToConstraints may contain sequences of the same value where the Excel constraint is 'bin' or 'int' and so has no rows
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

#If DEBUG_ Then
826       Debug.Assert (row >= 1 And row <= NumRows)
827       Debug.Assert (constraint >= 1 And constraint <= UBound(MappingRowsToConstraints))
#End If
828       If MappingRowsToConstraints(constraint) > row Then constraint = 1 ' Start searching from the beginning (TODO: Bisection search)
829       While row >= MappingRowsToConstraints(constraint + 1)
830           constraint = constraint + 1
831       Wend
832       instance = row - MappingRowsToConstraints(constraint) + 1
#If DEBUG_ Then
833       Debug.Assert (constraint >= 1 And constraint <= UBound(MappingRowsToConstraints) And row >= MappingRowsToConstraints(constraint) And row < MappingRowsToConstraints(constraint + 1))
#End If

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("COpenSolver", "GetConstraintFromRow") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Sub GetArrayPosnFromConstraintInstance(constraint As Long, instance As Long, i As Long, j As Long)
          ' Get the row (x) and constraint that goes with this LP row
          ' The incoming Constraint value will be either 1 or the constraint found on the last call to this
          ' We return the entry where MappingRowsToConstraints(Constraint) <= Row < MappingRowsToConstraints(Constraint+1)
          ' Note that MappingRowsToConstraints may contain sequences of the same value where the Excel constraint is 'bin' or 'int' and so has no rows
834       i = 1 + ((instance - 1) \ UBound(LHSOriginalValues(constraint), 2))
835       j = 1 + ((instance - 1) Mod UBound(LHSOriginalValues(constraint), 2))
End Sub

Function CheckConstraintValuesAreNumeric(constraint As Long, LHSValues As Variant, RHSValues As Variant, ErrorMsg As String) As Boolean
          ' Check that the LHS and RHS values are doubles (not vbError, or vbString, or ??)
          ' vbEmpty (blank) cell values are permitted as VBA quietly turns them into 0's in expressions
836       CheckConstraintValuesAreNumeric = False
          Dim i As Long, j As Long, instance As Long
          Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
          
          ' Check LHS first; this is always an array
837       instance = 0
838       For i = 1 To UBound(LHSValues, 1)
839           For j = 1 To UBound(LHSValues, 2)
840               instance = instance + 1
841               If VarType(LHSValues(i, j)) = vbEmpty Then
                      ' This is an empty cell; VBA quietly turns this into 0 when used in an expression. We need no special handling of this, and don't report an error.
842               ElseIf VarType(LHSValues(i, j)) = vbError Then
843                   GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
844                   ErrorMsg = "The cell '" & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "' in constraint '" & ConstraintSummary(constraint) & "' has a divide by zero error or has a problem with the function used ( eg #DIV/0! or #VALUE!). Please fix this and try again."
845                   Exit Function
846               ElseIf VarType(LHSValues(i, j)) <> vbDouble Then
                      ' It might be vbError, or vbString, or vbEmpty
847                   GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
848                   ErrorMsg = "The cell '" & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "' in constraint '" & ConstraintSummary(constraint) & "' has a value which is not numeric."
849                   On Error Resume Next
                      ' Try to show the incorrect value; just cary on with the old ErrorMsg if this gives an error
850                   ErrorMsg = "The cell '" & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "' in constraint '" & ConstraintSummary(constraint) & "' has the value '" & LHSValues(i, j) & "' which is not numeric."
851                   Exit Function
852               End If
853           Next j
854       Next i
          
855       Select Case RHSType(constraint)
              Case SolverInputType.SingleCellRange, SolverInputType.Formula, SolverInputType.constant
856               If VarType(RHSValues) = vbEmpty Then
                      ' This is an empty cell; VBA quietly turns this into 0 when used in an expression. We need no special handling of this, and don't report an error.
857               ElseIf VarType(RHSValues) = vbError Then
858                   GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
859                   ErrorMsg = "The right hand side in constraint '" & ConstraintSummary(constraint) & "' has a divide by zero error or has a problem with the function used ( eg #DIV/0! or #VALUE!). Please fix this and try again."
860                   Exit Function
861               ElseIf VarType(RHSValues) <> vbDouble Then
                      ' It might be vbError, or vbString, or anything else?
862                   GetConstraintInstanceData constraint, 1, LHSCellRange, RHSCellRange, RHSstring
863                   ErrorMsg = "The right hand side in constraint '" & ConstraintSummary(constraint) & "' has a value which is not numeric."
                      ' Try to show the incorrect value; just cary on with the old ErrorMsg if this gives an error
864                   On Error Resume Next
865                   ErrorMsg = "The right hand side in constraint '" & ConstraintSummary(constraint) & "' has the value '" & RHSValues & "' which is not numeric."
866                   Exit Function
867               End If
868           Case SolverInputType.MultiCellRange
869               instance = 0
870               For i = 1 To UBound(RHSValues, 1)
871                   For j = 1 To UBound(RHSValues, 2)
872                       instance = instance + 1
873                       If VarType(RHSValues(i, j)) = vbEmpty Then
                              ' This is an empty cell; VBA quietly turns this into 0 when used in an expression. We need no special handling of this, and don't report an error.
874                       ElseIf VarType(RHSValues) = vbError Then
875                           GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
876                           ErrorMsg = "The right hand side in constraint '" & ConstraintSummary(constraint) & "' has a divide by zero error or has a problem with the function used ( eg #DIV/0! or #VALUE!). Please fix this and try again."
877                           Exit Function
878                       ElseIf VarType(RHSValues(i, j)) <> vbDouble Then
                              ' It might be vbError, or vbString, or vbEmpty
879                           GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
880                           ErrorMsg = "The cell '" & RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "' in constraint '" & ConstraintSummary(constraint) & "' has a value which is not numeric."
881                           On Error Resume Next
882                           ErrorMsg = "The cell '" & RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "' in constraint '" & ConstraintSummary(constraint) & "' has the value '" & RHSValues(i, j) & "' which is not numeric."
883                           Exit Function
884                       End If
885                   Next j
886               Next i
887       End Select
888       CheckConstraintValuesAreNumeric = True
End Function

' Get the current values off the spreadsheet for the given constraint. LHS is always returned as a 2D array. RHS may be a single or array value
Sub GetCurrentConstraintValues(constraint As Long, LHSValues As Variant, RHSValues As Variant)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

889       If LHSType(constraint) = SolverInputType.SingleCellRange Then
              ' Create a 1x1 array
890           LHSValues = Create1x1Array(LHSRange(constraint).Value2) ' LHSRange(Constraint).value
891       Else ' LHSConstraintType = SolverInputType.MultiCellRange
892           LHSValues = LHSRange(constraint).Value2    ' Copy all the values into a variant, which will then be an array of doubles; see http://support.microsoft.com/kb/213798
893       End If

894       If RHSType(constraint) = SolverInputType.SingleCellRange Then
895           RHSValues = RHSRange(constraint).Value2  ' A variant containing just a single double
896       ElseIf RHSType(constraint) = SolverInputType.MultiCellRange Then
897           RHSValues = RHSRange(constraint).Value2   ' Copy all the values into a variant, which will then be an array of doubles (or an array of variants?)
898       ElseIf RHSType(constraint) = SolverInputType.Formula Then
899           RHSValues = ActiveSheet.Evaluate("=" & RHSFormula(constraint)) ' Will not throw an error, but may put an error value in RHSValues
900       Else        'if the rhs value is a constant
901           RHSValues = RHSOriginalValues(constraint)
902       End If

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("COpenSolver", "GetCurrentConstraintValues") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

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
918           Case SolverInputType.Formula
919               Set RHSCellRange = Nothing
920               RHSstring = RHSFormula(constraint)
921           Case SolverInputType.constant
922               Set RHSCellRange = Nothing
923               RHSstring = CStr(RHSOriginalValues(constraint))
924       End Select

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("COpenSolver", "GetConstraintInstanceData") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Function GetDecisionVariableValuesOffSheet() As Double()
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          ' Return an array containing the current values of the decision variables on the sheet.
          ' This needs improving for speed
          ' Note that the decision variables are defined by a range that may have multiple areas
          '      but must be all on one sheet (even if not the ActiveSheet)
          Dim Vars() As Double, var As Long
925       ReDim Vars(numVars)
926       For var = 1 To numVars
927           Vars(var) = AdjustableCells.Worksheet.Range(VarNames(var)).Value2
928       Next var
929       GetDecisionVariableValuesOffSheet = Vars

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("COpenSolver", "GetDecisionVariableValuesOffSheet") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function
    
Function CalcObjFnValue(DecisionVariableValues() As Double) As Double
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim sum As Double, var As Long
930       sum = ObjectiveFunctionConstant ' The value found for the obj function when all decision variables are 0
931       For var = 1 To numVars
932           sum = sum + (CostCoeffs(var) * DecisionVariableValues(var))
933       Next var
934       CalcObjFnValue = sum

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("COpenSolver", "CalcObjFnValue") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Sub BuildModelFromSolverData(Optional LinearityCheckOffset As Double = 0, _
                             Optional ShouldMinimiseUserInteraction As Boolean = False, _
                             Optional ShouldSolveRelaxation As Boolean = False)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

935       SolveStatus = OpenSolverResult.Unsolved
936       SolveStatusString = "Unsolved"
937       SolveStatusComment = ""
940       SolutionWasLoaded = False
          
941       ModelStatus = Unitialized

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
          
946       CheckWorksheetAvailable SuppressDialogs:=True, ThrowError:=True ' Will throw if there is no workbook or worksheet

947       Application.EnableCancelKey = xlErrorHandler
          
958       Set book = ActiveWorkbook
959       RawWorkBookName = book.Name
950       Set sheet = book.ActiveSheet
956       RawSheetName = sheet.Name
957       sheetName = EscapeSheetName(sheet)
          
        
          '----------------------------------------------------------------------------------------------------------
          'Find the solver
960       Set Solver = CreateSolver(GetChosenSolver(book, sheet))
          OpenSolverErrorHandler.SolverInUse = Solver.ShortName

          '--------------------------------------------------------------------------
          ' ADJUSTABLE CELLS SETUP and Setting of Initial Decision Var Cell values
          '--------------------------------------------------------------------------

          ' Raises an error if we can't determine a range or the name is missing.
          ' Remove any overlap
980       Set AdjustableCells = GetDecisionVariablesNoOverlap(book, sheet)
          ' Handle any merged ranges
          Dim BadCell As Range
981       If Not CheckRangeContainsNoAmbiguousMergedCells(AdjustableCells, BadCell) Then
983           Err.Raise Number:=OpenSolver_BuildError, Description:="This model contains a decision variable " & BadCell.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & " that is inaccessible as it is within the merged range " & BadCell.MergeArea.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & ". Please redefine the decision variable cells, and try again."
984           GoTo ExitSub
985       End If
          ' Check we have any cells after all this
986       numVars = AdjustableCells.Count
987       If numVars < 1 Then
989           Err.Raise Number:=OpenSolver_BuildError, Description:="There are no decision variable cells in this model, and so the model cannot be solved."
990       End If

          ' Get names for all the variables, creating names that distinguish cells on different sheets
991       ReDim pVarNames(numVars)
992       ReDim pCostCoeffs(numVars) As Double
          Dim AdjCell As Range, i As Long
996       i = 0
997       For Each AdjCell In AdjustableCells
998           i = i + 1
999           VarNames(i) = AdjCell.Address(RowAbsolute:=False, ColumnAbsolute:=False) ' eg A1
1003      Next AdjCell
          
          ' Set the cell values to catch protected sheets, and to set up for linear model building
1004      If SolverLinearity(Solver) = Linear Then
1005          AdjustableCells.Value2 = LinearityCheckOffset
          Else
              TestCellsForWriting AdjustableCells
1006      End If
          
          '--------------------------------------------------------------------------
          ' INITIAL SHEET CALCULATION (using new decision var values set just above)
          '--------------------------------------------------------------------------
          ' We expect, after this, that our calculation state will be xlDone. However, for large s/sheets, the
          ' state may be xlCalculating or xlPending even after a calculation.
          ' if this happens, then we cannot be sure that Application.Calculate will work, so we warn the user.
          ' TODO: Currently we require the calculation to complete before we continue. Should we make this more flexible?
1007      If Not ForceCalculate("Warning: The initial worksheet calculation did not complete, and so the model may not be generated correctly. Would you like to try again?", MinimiseUserInteraction) Then
1008          On Error GoTo ErrorHandler
1009          Err.Raise Number:=OpenSolver_BuildError, Description:="Calculation failed. The model has not been built so cannot be solved."
1010      End If
          '--------------------------------------------------------------------------
          ' OBJECTIVE SETUP
          '--------------------------------------------------------------------------

1011      ObjectiveSense = GetObjectiveSense(book, sheet)
1013      ObjectiveTargetValue = GetObjectiveTargetValue(book, sheet)

          Set ObjRange = GetObjectiveFunctionCellWithValidation(book, sheet)
1016      If Not ObjRange Is Nothing Then
              ' Store the value currently in the objective cell
1038          ObjRange.Calculate ' This does not always seem to work in that it doesn't track back predecessors; we now do a full re-calc above anyway
1039          objValue = ObjRange.Value2
1040          ObjectiveFunctionConstant = ObjRange.Value2
          End If

ModelOptions:

1041      AssumeNonNegativeVars = GetNonNegativity()

1045      If Not BuildModelOptions(ShouldSolveRelaxation, ShouldMinimiseUserInteraction) Then GoTo ExitSub

          Dim SolveOptions As SolveOptionsType
1557      GetSolveOptions sheet, SolveOptions

          ShowIterationResults = SolveOptions.ShowIterationResults
          
          ' Collate all solver parameters (ours and user-defined)
          Set SolverParameters = New Dictionary
1563      PopulateSolverParameters Solver, sheet, SolverParameters, SolveOptions
          
          '--------------------------------------------------------------------------
          ' CONSTRAINTS SETUP
          '--------------------------------------------------------------------------
          ' Count the correct number of constraints - not necessarily the number of
          ' rows though.
1048      NumConstraints = GetNumConstraints(book, sheet)
1049      UpdateStatusBar "OpenSolver: Setting Up Problem... " & numVars & " vars, " & NumConstraints & " Solver constraints", True
          'Set up arrays to be the correct size. We allow no constraints!
1050      If NumConstraints > 0 Then
1051          ReDim pLHSOriginalValues(NumConstraints) As Variant
1054          ReDim pLHSType(NumConstraints) As SolverInputType
1055          ReDim pLHSRange(NumConstraints) As Range
              ReDim pRELs(NumConstraints) As RelationConsts
1052          ReDim pRHSOriginalValues(NumConstraints) As Variant
1053          ReDim pRHSType(NumConstraints) As SolverInputType
1056          ReDim pRHSRange(NumConstraints) As Range
1057          ReDim pRHSFormula(NumConstraints) As String
1058          ReDim pConstraintSummary(NumConstraints) As String
1059          ReDim MappingRowsToConstraints(NumConstraints + 1) As Long
1060      End If
1061      Set VarLowerBounds = New Collection
1062      Set VarUpperBounds = New Collection
1063      ReDim pRowSetsBound(1)
          ' Reset number of rows in the LP model
1064      NumRows = 0
          Dim constraint As Long, row As Long
1065      row = 1
          ' Iterate through each solver constraint
1066      For constraint = 1 To NumConstraints
1067          UpdateStatusBar "OpenSolver: Setting Up Constraints... " & numVars & " vars, " & constraint & "/" & NumConstraints & " Solver constraints"
              
              '**********************************************************************
              ' Note: Solver enforces the following requirements.
              ' The LHS mut be a range with one or more cells
              ' The RHS can be either:
              '  - A single-cell range (=A4)
              '  - A multi-cell range of the same size as the LHS (=A4:B5)
              '  - A single constant value (eg =2)
              '  - A formula returning a single value (eg =sin(A4))
              '**********************************************************************
              
              '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
              ' LEFT HAND SIDE
              '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
              Dim rangeLHS As Range, sRefersToLHS As String
              Set rangeLHS = GetConstraintLhs(constraint, book, sheet, sRefersToLHS)
              
              '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
              ' RELATIONSHIP
              '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
              Dim rel As Long
              ' Either <=, = or >=, or int or bin (or 'alldif' in Excel 2010)
1081          rel = GetConstraintRel(constraint, book, sheet)
              RELs(constraint) = rel
                      
              Dim RHSCount As Long, RowCount As Long
1082          RowCount = rangeLHS.Count
              
1083          Select Case rel
              Case RelationINT, RelationBIN
                  ' Make the LHS variables integer or binary
                  ' Check that they haven't set as integer or binary something that
                  ' is not a decision variable...
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
1092                  Else 'if the relation is binary
1093                      Set BinaryCellsRange = ProperUnion(BinaryCellsRange, rangeLHS)
1094                  End If
1095              Else
                      ' The cells are not all decision variables
1096                  On Error GoTo ErrorHandler
1097                  Err.Raise Number:=OpenSolver_BuildError, Description:="Unable to find a cell specified as bin or int in the decision variables."
1098              End If
                  ' This constraint does not relate to any new rows, so mark it
                  ' with a dummy entry. NumRows+1 = same value as we get for the
                  ' next entry, so this refers to no rows.
1099              MappingRowsToConstraints(constraint) = NumRows + 1
              
              ' If the constraint is not int or bin...
              '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
              ' RIGHT HAND SIDE
              '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
1100          Case RelationGE, RelationLE, RelationEQ
                  ' ... then the constraint is a full equation with a RHS
                  Dim valRHS As Double, rangeRHS As Range, sRefersToRHS As String, RHSRefersToFormula As Boolean
1101              Set rangeRHS = GetConstraintRhs(constraint, sRefersToRHS, valRHS, RHSRefersToFormula, book, sheet)
                  
                  ' Track the constraints
                  Dim summary As String
1111              summary = ""
1112              summary = summary & StripWorksheetNameAndDollars(sRefersToLHS, sheet)
1113              summary = summary & " " & RelationEnumToString(rel) & " "
1114              summary = summary & StripWorksheetNameAndDollars(sRefersToRHS, sheet)
1115              ConstraintSummary(constraint) = summary

                  ' Check we have a compatible constraint system.
                  ' Solver should have checked this already, but just to make sure...
1120              RHSCount = 1
1121              If Not rangeRHS Is Nothing Then
                      RHSCount = rangeRHS.Count
                  End If
1122              If (RowCount <> 1 And RHSCount <> 1) And (RowCount <> RHSCount) Then
1123                  On Error GoTo ErrorHandler
1124                  Err.Raise Number:=OpenSolver_BuildError, Description:="The constraint '" & summary & "' has a different cell count on the left and the right. The model cannot be built."
1125              End If
                  
                  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                  ' STORE CONSTRAINT IN MEMORY
                  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                  ' Each constraint is stored as one row of LHSOriginalvalues and
                  ' RHSOriginalvalues which essentially become (variant) arrays of
                  ' variants. LHSType and RHSType tell us what is stored in each row
                  ' of LHSOriginalvalues and RHSOriginalvalues
                  
                  ' Left hand side:
1126              Set LHSRange(constraint) = rangeLHS
1127              If RowCount = 1 Then
1128                  LHSType(constraint) = SolverInputType.SingleCellRange
1129                  LHSOriginalValues(constraint) = Create1x1Array(rangeLHS.Value2)  ' Create a 1x1 array for the single cell range
1130              Else
1131                  LHSType(constraint) = SolverInputType.MultiCellRange
1132                  LHSOriginalValues(constraint) = rangeLHS.Value2 ' Copy the entire range over into an array of variants
1133              End If
                  
                  ' Right hand side:
                  ' If it is a range, it is much like LHS
1134              If Not rangeRHS Is Nothing Then
1135                  Set RHSRange(constraint) = rangeRHS
1136                  RHSType(constraint) = IIf(rangeRHS.Count = 1, SolverInputType.SingleCellRange, SolverInputType.MultiCellRange)
1137                  RHSOriginalValues(constraint) = rangeRHS.Value2  ' Copy the entire range over into an array of variants
                  ' If its a formula, store the formula internally as a string
1138              ElseIf RHSRefersToFormula Then
1139                  RHSFormula(constraint) = sRefersToRHS
1140                  RHSType(constraint) = SolverInputType.Formula
1141                  RHSOriginalValues(constraint) = ActiveSheet.Evaluate("=" & sRefersToRHS)
                  ' Its a constant
1142              Else
                      RHSFormula(constraint) = sRefersToRHS
1143                  RHSType(constraint) = SolverInputType.constant
1144                  RHSOriginalValues(constraint) = valRHS
1145              End If

                  ' Check if this constraint applies any bounds on decision variables
1146              If RowCount = 1 Then
                      ' Single LHS, single RHS
1147                  ApplyBounds row, rel, rangeLHS, rangeRHS, valRHS, RHSRefersToFormula
1148              Else
                      Dim CurRow As Long, RHSRowCount As Long
1149                  If rangeRHS Is Nothing Then
1150                      RHSRowCount = 1
1151                  Else
1152                      RHSRowCount = rangeRHS.Count
1153                  End If
                      
1154                  If RHSRowCount > 1 Then
                          ' multiple LHS, multiple RHS
1155                      For CurRow = 1 To RowCount
1156                          ApplyBounds row + CurRow - 1, rel, rangeLHS(CurRow), rangeRHS(CurRow), valRHS, RHSRefersToFormula
1157                      Next CurRow
1158                  Else
                          ' multiple LHS, single RHS
1159                      For CurRow = 1 To RowCount
1160                          ApplyBounds row + CurRow - 1, rel, rangeLHS(CurRow), rangeRHS, valRHS, RHSRefersToFormula
1161                      Next CurRow
1162                  End If
1163              End If
                  
                  ' Store relations for these rows, keep track of how many rows we've added
1164              ReDim Preserve pRelation(NumRows + rangeLHS.Count) As RelationConsts
1165              For i = 1 To rangeLHS.Count
1166                  Relation(row + i - 1) = rel
1167              Next i
1168              row = row + rangeLHS.Count
                  ' Store the first LP row associated with this Excel constraint
1169              MappingRowsToConstraints(constraint) = NumRows + 1
1170              NumRows = NumRows + RowCount
                  
                  ' Check that all our values are in fact numbers
                  Dim ErrorMsg As String
1171              If Not CheckConstraintValuesAreNumeric(constraint, LHSOriginalValues(constraint), RHSOriginalValues(constraint), ErrorMsg) Then
1172                  On Error GoTo ErrorHandler
1173                  Err.Raise Number:=OpenSolver_BuildError, Description:=ErrorMsg
1174              End If
              
              Case Else
                  ' Don't allow 'alldiff' from Excel 2010 or any other wierd things
1116              On Error GoTo ErrorHandler
1118              Err.Raise Number:=OpenSolver_BuildError, Description:="Unrecognised relationship for constraint " & ConstraintSummary(constraint)
1175          End Select
1176      Next constraint

          ' Binary trumps integer, so remove all overlap
          Set IntegerCellsRange = SetDifference(IntegerCellsRange, BinaryCellsRange)

1177      If NumRows > 0 Then
1178          ReDim Preserve pRowSetsBound(NumRows)
1179      End If
          
          'Set up extra model info depending on the solver
1180      Select Case Solver.ModelType
          Case Diff
              ' BUILD THE A MATRIX
1181          If NumRows > 0 Then
                  ReDim Preserve pSparseA(NumRows) As CIndexedCoeffs
                  ReDim pRHS(NumRows) As Double
              End If
1182          For row = 1 To NumRows
1183              Set SparseA(row) = New CIndexedCoeffs
1184          Next row
1186          If Not BuildConstraintMatrix(ObjRange Is Nothing, pSparseA) Then GoTo ExitSub
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
1230          If Not RHS.HasFormula Then
                  ' The RHS range is a constant - apply a bound
1231              UpdateBound row, LHS.Address, CDbl(RHS), rel
1232          Else
                  ' The RHS range is a formula, check if this depends on the decision vars
                  ' Will raise an error is no precedents are found (e.g. if they are on another sheet
1234              If Intersect(RHS.Precedents, AdjustableCells) Is Nothing Then
1235                  UpdateBound row, LHS.Address, CDbl(RHS), rel
1236              End If
1237          End If
1238      Else
              ' The RHS is a numeric constant - apply a bound
1239          UpdateBound row, LHS.Address, RHSValue, rel
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
          
          ' Only GE constraints can provide lower bounds
1246      If rel = RelationGE Then
              ' Remove existing lower bound on variable if our new one is better
1247          If TestKeyExists(VarLowerBounds, VarName) Then
1248              If bound > VarLowerBounds(VarName) Then
1249                  VarLowerBounds.Remove VarName
1250              End If
1251          End If
              
1252          VarLowerBounds.Add bound, VarName
          ' Only LE constraints can provide upper bounds
1253      ElseIf rel = RelationLE Then
              ' Remove existing upper bound on variable if our new one is better
1254          If TestKeyExists(VarUpperBounds, VarName) Then
1255              If bound < VarUpperBounds(VarName) Then
1256                  VarUpperBounds.Remove VarName
1257              End If
1258          End If
              
1259          VarUpperBounds.Add bound, VarName
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

Function BuildModelOptions(Optional ShouldSolveRelaxation As Boolean = False, Optional ShouldMinimiseUserInteraction As Boolean = False) As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          BuildModelOptions = False
          
          ' Ignore integrality check if the user is solving the relaxation
          If Not ShouldSolveRelaxation Then
              ' If we are in Excel 2010 or newer, and the user has set "Solve Relaxtion", then we report an error
              Dim IngoreIntConstraints As Boolean
1265          IngoreIntConstraints = False
1266          If Val(Application.Version) >= 14 And GetIgnoreIntegerConstraints(book, sheet) Then
                   If ShouldMinimiseUserInteraction Then GoTo RaiseError
                   If Not MsgBox("You have the 'Ignore integer constraints' option enabled in Solver, so OpenSolver cannot proceed with solving. You can either disable this option, or use the 'Solve Relaxation' option in the OpenSolver menu. Would you like OpenSolver to disable this option and proceed?", vbCritical + vbYesNo + vbDefaultButton1, "OpenSolver Warning") = vbNo Then
                       SetIgnoreIntegerConstraints False, book, sheet
                   Else
RaiseError:
                       ' Set UserCancelledError because the user has already confirmed the abort if not in silent mode
                       Err.Raise Number:=OpenSolver_UserCancelledError, Description:="You have the Solver option 'ignore integer constraints' turned on. Please turn this off in the Solver (not OpenSolver) options menu and then solve with OpenSolver again or choose OpenSolver's solve relaxed model option"
                   End If
1271          End If
          End If

1272      BuildModelOptions = True

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("COpenSolver", "BuildModelOptions") Then Resume
          RaiseError = True
          GoTo ExitSub
End Function

Function BuildConstraintMatrix(objIsMissing As Boolean, SparseA() As CIndexedCoeffs) As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

1275      BuildConstraintMatrix = False
          
          '--------------------------------------------------------------------------
          ' FILL A MATRIX BY CHANGING ADJUSTABLE CELLS
          '--------------------------------------------------------------------------
          ' Incrementing each decision variable by one to see what constraints are
          ' dependent on it and by what amount
          Dim var As Long
          Dim AdjCell As Range
          
          ' Add constants to Target
1276      If ObjectiveSense = TargetObjective Then
              ' Target value needs to be incremented by any constants in the objective
1277          ObjectiveTargetValue = ObjectiveTargetValue - objValue
1278      End If
          
1279      var = 0
1280      For Each AdjCell In AdjustableCells
              ' Count how many variables we've processed, let the user know
              ' that the program hasn't crashed!
1281          var = var + 1
1282          UpdateStatusBar "OpenSolver: Setting Up Problem... " & var & "/" & numVars & " vars, " & NumRows & " rows."

              ' Increment the value of this cell to find its coefficients
              ' Because its all linear, we can simply check the change in the objective/constraints
1283          AdjCell.Value2 = AdjCell.Value2 + 1

              ' Recalculate the spreadsheet
1284          If Not ForceCalculate("Warning: The worksheet calculation did not complete, and so the model may not be generated correctly. Would you like to retry?") Then GoTo ExitSub

              ' Easy part first - the objective function value change
1285          If Not objIsMissing Then
1286              CostCoeffs(var) = ObjRange.Value2 - objValue
1287          End If

              ' Now the hard (slow) bit - the constraint changes
              Dim CurrentLHSValues As Variant
              Dim CurrentRHSValues As Variant
              Dim row As Long, constraint As Long
1288          row = 1
1289          For constraint = 1 To NumConstraints

                  ' Check to see what is different and add rows to SparseA
1290              If Not LHSRange(constraint) Is Nothing Then ' skip Binary and Integer constraints

                      Dim OriginalLHSValues As Variant
                      Dim OriginalRHSValues As Variant
1291                  OriginalLHSValues = LHSOriginalValues(constraint)
1292                  OriginalRHSValues = RHSOriginalValues(constraint)

                      ' Get current value(s) for LHS and RHS of this constraint off the sheet. LHS is always an array (even if 1x1)
1293                  GetCurrentConstraintValues constraint, CurrentLHSValues, CurrentRHSValues

                      Dim coeff As Double, i As Long, j As Long
1294                  For i = 1 To UBound(OriginalLHSValues, 1) ' number of rows
1295                      For j = 1 To UBound(OriginalLHSValues, 2) ' number of columns
1296                          If RHSType(constraint) = SolverInputType.MultiCellRange Then
                                  '-----------------------------------------------------------
                                  'Making it work for column LHS with row RHS and vice versa
                                  '-----------------------------------------------------------
1297                              If UBound(OriginalLHSValues, 1) = UBound(OriginalRHSValues, 1) Then
1298                                  coeff = (CurrentLHSValues(i, j) - OriginalLHSValues(i, j)) - (CurrentRHSValues(i, j) - OriginalRHSValues(i, j))
1299                              Else
1300                                  coeff = (CurrentLHSValues(i, j) - OriginalLHSValues(i, j)) - (CurrentRHSValues(j, i) - OriginalRHSValues(j, i))
1301                              End If
1302                          ElseIf RHSType(constraint) = SolverInputType.Formula Or RHSType(constraint) = SolverInputType.SingleCellRange Then
1303                              coeff = (CurrentLHSValues(i, j) - OriginalLHSValues(i, j)) - (CurrentRHSValues - OriginalRHSValues)
1304                          Else    ' if the rhs is a constant value
1305                              coeff = (CurrentLHSValues(i, j) - OriginalLHSValues(i, j))
1306                          End If
                              ' Ignore zero (or near-zero) coefficients
1307                          If Abs(coeff) > EPSILON Then
1308                              SparseA(row).Add var, coeff
1309                          End If
1310                          row = row + 1
1311                      Next j
1312                  Next i
                      ' End If
1313              End If
1314          Next constraint
1315          AdjCell.Value2 = AdjCell.Value2 - 1
1316      Next AdjCell
          
1317      If (NumConstraints > 0) Then
1318          MappingRowsToConstraints(NumConstraints + 1) = NumRows + 1  ' An 'end of data' entry
1319      End If

          'Create the rhs vector
          ' TODO Handle errors in the original spreadsheet values
1320      row = 1
1321      For constraint = 1 To NumConstraints
1322          If Not LHSRange(constraint) Is Nothing Then
1323              If RHSType(constraint) = SolverInputType.MultiCellRange Then 'does the constraint refer to more than one cell
1324                  For i = 1 To UBound(LHSOriginalValues(constraint), 1)
1325                      For j = 1 To UBound(LHSOriginalValues(constraint), 2)
1326                          If UBound(LHSOriginalValues(constraint), 1) = UBound(RHSOriginalValues(constraint), 1) Then
1327                              RHS(row) = RHSOriginalValues(constraint)(i, j) - LHSOriginalValues(constraint)(i, j)
1328                          Else
1329                              RHS(row) = RHSOriginalValues(constraint)(j, i) - LHSOriginalValues(constraint)(i, j)
1330                          End If
1331                          row = row + 1
1332                      Next j
1333                  Next i
1334              Else
1335                  For i = 1 To UBound(LHSOriginalValues(constraint), 1)
1336                      For j = 1 To UBound(LHSOriginalValues(constraint), 2)
1337                          RHS(row) = RHSOriginalValues(constraint) - LHSOriginalValues(constraint)(i, j)
1338                          row = row + 1
1339                      Next j
1340                  Next i
1341              End If
1342          End If
1343      Next constraint

1344      BuildConstraintMatrix = True
          
ExitSub:
          Application.StatusBar = False
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("COpenSolver", "BuildConstraintMatrix") Then Resume
          RaiseError = True
          GoTo ExitSub
End Function

Sub ViewSparseA()
                Dim i As Long, j As Long, strA As String
1367            For i = 1 To UBound(pSparseA)
1368                For j = 1 To numVars
1369                    strA = strA & CStr(SparseA(i).Coefficient(j)) & "    "
1370                Next j
1371                strA = strA & vbNewLine
1372            Next i
1373            MsgBox (strA)
End Sub


Function CheckConstantConstraintIsSatisfied(row As Long, constraint As Long, instance As Long, i As Long, j As Long) As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim LHSValue As Double, RHSValue As Double
1886      GetArrayPosnFromConstraintInstance constraint, instance, i, j
1887      LHSValue = LHSOriginalValues(constraint)(i, j)
1888      If RHSType(constraint) <> MultiCellRange Then
1889          RHSValue = RHSOriginalValues(constraint)
1890      Else
1891          RHSValue = RHSOriginalValues(constraint)(i, j)
1892      End If
          ' GetExcelConstraintInfoFromRow Constraint, Row, LHSValue, RHSValue, Probleminstance, sLHS, sRHS
1893      If (Relation(row) = RelationGE And LHSValue < RHSValue - EPSILON) _
             Or (Relation(row) = RelationEQ And Abs(LHSValue - RHSValue) > EPSILON) _
             Or (Relation(row) = RelationLE And LHSValue > RHSValue + EPSILON) Then
              ' AJM 28.11.2010 We cannot use any IIF() that references LHSRange(row).Address because all arguments are always evaluated, and LHSRange(row).Address is not defined and throws an error
              'sLHS = LHSRange(j).Address & " = " & lhsvalue
              Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
1894          GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
              ' If the RHS is a range, we show its address; if not, RHSString contains the RHS's constant or formula
1895          If Not RHSCellRange Is Nothing Then RHSstring = RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
1896          SolveStatus = OpenSolverResult.Infeasible
1897          SolveStatusString = "Infeasible"
1898          SolveStatusComment = "The model contains a constraint '" & ConstraintSummary(constraint) & "' for which instance " _
                     & instance & " does not depend on the decision variables and is not satisfied." & vbCrLf _
                    & "Constraint specifies: " _
                     & "LHS=" & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "=" & LHSValue _
                    & " " & RelationEnumToString(Relation(row)) & " " _
                    & "RHS=" & StripWorksheetNameAndDollars(RHSstring, sheet) & "=" & RHSValue
1899          Close #1
1900          CheckConstantConstraintIsSatisfied = False
1901          GoTo ExitFunction
1902      End If
1903      CheckConstantConstraintIsSatisfied = True

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("COpenSolver", "CheckConstantConstraintIsSatisfied") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Sub ReportAnySolutionSubOptimality()
          ' Call this after SolveModel to report to the user what happened. A dialog is shown if the optimal result was not found
          Select Case SolveStatus
              Case OpenSolverResult.Optimal                ' We are optimal
              Case OpenSolverResult.Pending                ' The solve is scheduled
              Case OpenSolverResult.NotLinear              ' Already shown non-linearity details
              Case OpenSolverResult.AbortedThruUserAction  ' Already shown escape dialog
          
              Case Else
1909              MsgBox "OpenSolver could not find an optimal solution, and reported:" & vbCrLf & _
                         SolveStatusString & vbCrLf & vbCrLf & _
                         IIf(SolveStatusComment = "", "", SolveStatusComment & vbCrLf & vbCrLf) & _
                         IIf(SolutionWasLoaded, "The solution generated has been loaded into the spreadsheet.", "No solution was available to load into the spreadsheet.")
1910      End Select
End Sub

Sub ClearQuickSolve()
1992      Set QuickSolve = Nothing
End Sub

Function InitializeQuickSolve() As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

1993      InitializeQuickSolve = False
1994      If ModelStatus <> Built Then
1995          MsgBox "QuickSolve Initialisation Failed: The model cannot be solved as it has not yet been built.", , "OpenSolver" & sOpenSolverVersion & " Error"
1996          GoTo ExitFunction
1997      End If

1998      Set QuickSolve = New CQuickSolve
1999      If QuickSolve.AnalyseParameters(Me, sheet, NumConstraints, pRHS, pLHSRange, pRHSRange, pRHSFormula, pRHSType, pLHSOriginalValues, pRHSOriginalValues) Then
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

Function CanDoQuickSolveForActiveSheet() As Boolean
          ' Confirms that we can do a quicksolve for the active sheet because that is the sheet we are set up for
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

2004      CanDoQuickSolveForActiveSheet = False
2005      If QuickSolve Is Nothing Then
2006          MsgBox "Error: The quick solve data has not been initialized. Please choose Intialize Quick Solve, and then try again.", , "OpenSolver" & sOpenSolverVersion & " Error"
2007      ElseIf (ActiveWorkbook.Name <> RawWorkBookName) Or (ActiveSheet.Name <> RawSheetName) Then
2008          MsgBox "Error: The quick solve data has not been initialized for this worksheet. Please choose Intialize Quick Solve, and then try again.", , "OpenSolver" & sOpenSolverVersion & " Error"
2009      Else
2010          CanDoQuickSolveForActiveSheet = True
2011      End If

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("COpenSolver", "CanDoQuickSolveForActiveSheet") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Sub DoQuickSolve(Optional MinimiseUserInteraction As Boolean = False)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
2013      If QuickSolve Is Nothing Then
2014          MsgBox "Error: The quick solve data has not been initialized. Please choose Intialize Quick Solve, and then try again.", , "OpenSolver" & sOpenSolverVersion & " Error"
2015      Else
2016          If QuickSolve.UpdateModelRHS(sheet, pRHS) Then
2017              SolveModel Me, False, MinimiseUserInteraction
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
          ReDim pFinalVarValue(numVars)
1435      ReDim pVarCell(numVars)
1436      If bGetDuals Then
1437          ReDim pFinalValue(NumRows)
1438          ReDim pShadowPrice(NumRows)
1439          ReDim pIncreaseCon(NumRows)
1440          ReDim pDecreaseCon(NumRows)
1441          ReDim pReducedCosts(numVars)
1442          ReDim pIncreaseVar(numVars)
1443          ReDim pDecreaseVar(numVars)
1444      End If
End Sub

Public Sub LoadResultsToSheet()
   If SolutionWasLoaded Then
        AdjustableCells.Value2 = 0
        Dim i As Long
        For i = 1 To numVars
            If Len(VarCell(i)) <> 0 Then
                AdjustableCells.Worksheet.Range(VarCell(i)).Value2 = FinalVarValue(i)
            End If
        Next i
    End If
End Sub
