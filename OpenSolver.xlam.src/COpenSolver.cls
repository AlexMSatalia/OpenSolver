VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COpenSolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' v0.2: Switched to APplication.Calculation = manual; it is twice as fast. Looping thru the LHS and RHS ranges is an insignificant time when compared to the calculation time
'       Eg, Run time with no LHS and RHS range loopups is 4.2s, this goes to 4.4 or 4.5 when we loop thru the LHS and RHS ranges
'       Note: Reading cell values one by one is very slow.
'       Instead, see:
'          http://www.xtremevbtalk.com/showthread.php?t=296858
'          http://www.avdf.com/apr98/art_ot003.html
'          http://www.food-info.net/uk/e/e173.htm - very good info on writing fast code
'          http://blogs.msdn.com/excel/archive/2009/03/12/excel-vba-performance-coding-best-practices.aspx - fast coding
'          http://msdn.microsoft.com/en-us/library/aa730921.aspx - microsoft info on Excel 2007 what's new
'          http://support.microsoft.com/kb/153090/EN-US/ - pass an Excel array to VB
'          http://support.microsoft.com/kb/177991 - limitations when passing arrays to sheets

#Const DEBUG_ = True

Option Explicit
Option Base 1

Public sheetName As String ' Sheetname in quotes as used for prefixing
Dim RawSheetName As String ' Just the worksheet name
Dim RawWorkBookName As String ' Just the workbook name
Dim book As Workbook
Public sheet As Worksheet

' Whether to show user alerts
Dim MinimiseUserInteraction As Boolean

' Data for our model. We work thru the original Excel 'constraints' and form model 'rows' from these
Public AdjustableCells As Range    ' A range we form containing all the adjustable (ie decision) variables
Public BinaryCellsRange As Range   ' A range we form containing all the adjustable (ie decision) variables that are binary
Public IntegerCellsRange As Range  ' A range we form containing all the adjustable (ie decision) variables that are integer (which may include binary variables)
Public AssumeNonNegativeVars As Boolean   ' Excel stores ...!solver_neg=1 if Assume Linear Model is turned on, 2 if off (NB: Not 0!)

' For the LHS and RHS of each Excel constraint, we store its type, and its LHS and RHS
Enum SolverInputType
    SingleCellRange = 1 ' Valid for a LHS and a RHS
    MultiCellRange = 2  ' Valid for a LHS and a RHS
    Formula = 3         ' Valid for a RHS only
    constant = 4        ' Valid for a RHS only
End Enum
Dim LHSType() As SolverInputType
Dim RHSType() As SolverInputType
Dim LHSRange() As Range, RHSRange() As Range
Dim RHSFormula() As String
Dim ConstraintSummary() As String

' The original values of the LHS and RHS of the Excel constraints before we start perturbing the adjustable cells;
' each entry is stored as a variant that could be a double or an array
Dim LHSOriginalValues() As Variant
Dim RHSOriginalValues() As Variant

Public ObjRange As Range

' Details for the model that we build.
Public numVars As Long, NumIntVars As Long, NumBinVars As Long, NumConstraints As Long
Public NumRows As Long
Public objValue As Double  ' The constant value of the objective
Public ObjectiveSense As Long
Public ObjectiveTargetValue As Double  ' Used when the user wants the objective to achieve some stated target value
Dim VarNames() As String
Dim AdjCellsSheetIndex As Long    ' The sheet index containing variables with no index specified in the CBC result file
Dim pCostCoeffs() As Double, ObjectiveFunctionConstant As Double
Dim RHS() As Double ' The RHS values for each row in the LP
Dim Relation() As Long ' For each row, we remember its relation (=, <=, >=)
Public VarLowerBounds As Collection
Public VarUpperBounds As Collection

'Details for the chosen solver
Public Solver As String 'Name of solver we will be using
Public ExternalSolverPathName As String 'Path to solver
Public ModelFilePathName As String ' Path to model file
Dim SolutionFilePathName As String ' Path to model file

' Details we need for non-linear solver
Dim VariableBounds() As Boolean
Dim StartingPosition() As Double
Dim SolveRelaxedModel As Boolean

' Details for sensitivity analysis
Public bGetDuals As Boolean ' Whether to get duals
Public rConstraintList As Range
Dim DualsOnNewSheet As Boolean
Dim DualsOnSameSheet As Boolean

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
Dim SparseA() As CIndexedCoeffs

' We remember how to turn a row index in the LP back into a constraint index in the Excel list of constraints
' This array stores the first row in the LP associated with each Excel constraint
Dim MappingRowsToConstraints() As Long

' Our result as an enum and a string
Public SolveStatus As OpenSolverResult ' See Enum in OpenSolverModule; Optimal,Infeasible, etc
Public SolveStatusString As String    ' Eg Infeasible, Opimal etc
Public SolveStatusComment As String ' Eg constraint 1 cannot be satisified
Public SolutionWasLoaded As Boolean ' True if we loaded a solution into the spreadsheet

Public ModelStatus As Long

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

Public Property Get VarNamesP(Index As Long) As String
820       VarNamesP = VarNames(Index)
End Property

Public Property Get LHSTypeP(Index As Long) As String
821       LHSTypeP = LHSType(Index)
End Property

Public Property Get RHSTypeP(Index As Long) As String
822       RHSTypeP = RHSType(Index)
End Property

Public Property Get RelationP(Index As Long) As Long
823       RelationP = Relation(Index)
End Property
' End of getters and setters
'===============================================================================

' See http://www.jkp-ads.com/OfficeMarketPlaceNM-EN.asp for a tool to view the Solver's hidden names
' See https://www.coin-or.org/OS/publications/OSICS2009.pdf for accessing a COIN-OR solver
' See http://ideas.repec.org/a/jss/jstsof/14i04.html Calling LPSolve from Excel (by Samuel E. Buttrey)
' See http://peltiertech.com/Excel/SolverVBA.html for info on Solver and VBA
' http://www.sce.carleton.ca/faculty/chinneck/MPSreader/MPSobject.html for an MPS file reader/writer

' TODO: In an LP file, The maximum length for any name is 255. The maximum length of any line of input is 510.

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

Sub BuildModelFromSolverData(Optional LinearityCheckOffset As Double = 0, Optional MinimiseUserInteraction As Boolean = False, Optional SolveRelaxation As Boolean = False)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

935       SolveStatus = OpenSolverResult.Unsolved
936       SolveStatusString = "Unsolved"
937       SolveStatusComment = ""
940       SolutionWasLoaded = False
          
941       ModelStatus = ModelStatus_Unitialized

          ' Change to manual calculation mode
          Dim oldCalculationMode As Long
969       oldCalculationMode = Application.Calculation
970       Application.Calculation = xlCalculationManual
971       Application.ScreenUpdating = False
972       Application.Cursor = xlWait

          Dim StartTimeOfBuild As Single, StartDateTimeOfBuild As Date
          Dim EndTimeOfBuild As Single, EndDateTimeOfBuild As Date
944       StartTimeOfBuild = Timer    ' Warning: This goes back to zero at midnight!
945       StartDateTimeOfBuild = Now()
          
946       CheckWorksheetAvailable SuppressDialogs:=True, ThrowError:=True ' Will throw if there is no workbook or worksheet

947       Application.EnableCancelKey = xlErrorHandler
          
949       On Error Resume Next
950       Set sheet = ActiveWorkbook.ActiveSheet
          ' This error checking is not really needed as it also occurs in CheckWorkSheetAvailable
951       If Err.Number <> 0 Then
952           On Error GoTo ErrorHandler
953           Err.Raise Number:=OpenSolver_BuildError, Description:="Unable to access the active worksheet."
954       End If
955       On Error GoTo ErrorHandler
956       RawSheetName = sheet.Name
957       sheetName = EscapeSheetName(sheet)
          
958       Set book = ActiveWorkbook
959       RawWorkBookName = book.Name
        
          '----------------------------------------------------------------------------------------------------------
          'Find the users preferred solver
960       Solver = GetChosenSolver(book, sheet)

          '--------------------------------------------------------------------------
          ' ADJUSTABLE CELLS SETUP and Setting of Initial Decision Var Cell values
          '--------------------------------------------------------------------------

          ' Raises an error if we can't determine a range or the name is missing.
          ' Remove any overlap
979       On Error GoTo ErrorHandler
980       Set AdjustableCells = GetDecisionVariablesNoOverlap(book, sheet)
          ' Handle any merged ranges
          Dim BadCell As Range
981       If Not CheckRangeContainsNoAmbiguousMergedCells(AdjustableCells, BadCell) Then
982           On Error GoTo ErrorHandler
983           Err.Raise Number:=OpenSolver_BuildError, Description:="This model contains a decision variable " & BadCell.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & " that is inaccessible as it is within the merged range " & BadCell.MergeArea.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & ". Please redefine the decision variable cells, and try again."
984           GoTo ExitSub
985       End If
          ' Check we have any cells after all this
986       numVars = AdjustableCells.Count
987       If numVars < 1 Then
988           On Error GoTo ErrorHandler
989           Err.Raise Number:=OpenSolver_BuildError, Description:="There are no decision variable cells in this model, and so the model cannot be solved."
990       End If
          ' Find out which worksheet has the majority of our decision variables
          'Dim CellCount() As Long, area As Range
          'AdjCellsSheetIndex = 1
          'ReDim CellCount(ActiveWorkbook.Worksheets.Count)
          'For Each area In AdjustableCells.Areas
          '    CellCount(area.Worksheet.Index) = CellCount(area.Worksheet.Index) + area.Count
          '    If CellCount(area.Worksheet.Index) > CellCount(AdjCellsSheetIndex) Then AdjCellsSheetIndex = area.Worksheet.Index
          'Next area
          ' Get names for all the variables, creating names that distinguish cells on different sheets
991       ReDim VarNames(numVars)
992       ReDim pCostCoeffs(numVars) As Double
993       If Solver = "NOMAD" Then
994           ReDim StartingPosition(numVars)
995       End If
          Dim AdjCell As Range, i As Long
996       i = 0
997       For Each AdjCell In AdjustableCells
998           i = i + 1
999           VarNames(i) = AdjCell.Address(RowAbsolute:=False, ColumnAbsolute:=False) ' eg A1
1000          If Solver = "NOMAD" Then
1001              StartingPosition(i) = AdjCell.value
1002          End If
1003      Next AdjCell
          ' Set all the adjustable cells to zero (or a user-specified value) (For linear problems)
1004      If SolverType(Solver) = OpenSolver_SolverType.Linear Then
1005          AdjustableCells.Value2 = LinearityCheckOffset
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

          Dim succeeded As Boolean
1043      On Error GoTo ErrorHandler
1045      BuildModelOptions succeeded, SolveRelaxation, MinimiseUserInteraction
1046      If succeeded = False Then GoTo ExitSub
          
          '--------------------------------------------------------------------------
          ' CONSTRAINTS SETUP
          '--------------------------------------------------------------------------
          ' Count the correct number of constraints - not necessarily the number of
          ' rows though.
1048      NumConstraints = GetNumConstraints(sheet:=sheet)
1049      UpdateStatusBar "OpenSolver: Setting Up Problem... " & numVars & " vars, " & NumConstraints & " Solver constraints", True
          'Set up arrays to be the correct size. We allow no constraints!
1050      If NumConstraints > 0 Then
1051          ReDim LHSOriginalValues(NumConstraints) As Variant
1052          ReDim RHSOriginalValues(NumConstraints) As Variant
1053          ReDim RHSType(NumConstraints) As SolverInputType
1054          ReDim LHSType(NumConstraints) As SolverInputType
1055          ReDim LHSRange(NumConstraints) As Range
1056          ReDim RHSRange(NumConstraints) As Range
1057          ReDim RHSFormula(NumConstraints) As String
1058          ReDim ConstraintSummary(NumConstraints) As String
1059          ReDim MappingRowsToConstraints(NumConstraints + 1) As Long
1060      End If
1061      Set VarLowerBounds = New Collection
1062      Set VarUpperBounds = New Collection
1063      ReDim VariableBounds(1)
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
              Dim rangeLHS As Range
              Set rangeLHS = GetConstraintLhs(constraint, book, sheet)
              
              '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
              ' RELATIONSHIP
              '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
              Dim rel As Long
              ' Either <=, = or >=, or int or bin (or 'alldif' in Excel 2010)
1081          rel = GetConstraintRel(constraint, book, sheet)
                      
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
1112              summary = summary & Mid(Names(sheetName & "solver_lhs" & constraint), 2)
1113              summary = summary & " " & SolverRelationAsString(rel) & " "
1114              summary = summary & Mid(Names(sheetName & "solver_rhs" & constraint), 2)
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
1164              ReDim Preserve Relation(NumRows + rangeLHS.Count) As Long
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
1178          ReDim Preserve VariableBounds(NumRows)
1179      End If
          
          'Set up the constraint matrix for solving linear problems
1180      If SolverType(Solver) = OpenSolver_SolverType.Linear Then
              '--------------------------------------------------------------------------
              ' BUILD THE A MATRIX
              '--------------------------------------------------------------------------
              ' Initialise a sparse matrix
1181          If NumRows > 0 Then ReDim Preserve SparseA(NumRows) As CIndexedCoeffs
1182          For row = 1 To NumRows
1183              Set SparseA(row) = New CIndexedCoeffs
1184          Next row
1185          If NumRows > 0 Then ReDim RHS(NumRows) As Double
1186          BuildConstraintMatrix ObjRange Is Nothing, SparseA, succeeded
1187          If succeeded = False Then GoTo ExitSub
1188      End If

1189      ModelStatus = ModelStatus_Built
1190      EndTimeOfBuild = Timer
1191      EndDateTimeOfBuild = Now()
1192      SecondsTakenToBuildModel = DateDiff("s", StartDateTimeOfBuild, EndDateTimeOfBuild)
1193      TimeTakenToBuildModel = EndTimeOfBuild - StartTimeOfBuild

ExitSub:
          Application.Cursor = xlDefault
          Application.StatusBar = False
          Application.ScreenUpdating = True
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
1220      If Not SolverUsesUpperBounds(Solver) And rel = RelationLE Then GoTo ExitSub
           
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
1261      If UBound(VariableBounds) < row Then ReDim Preserve VariableBounds(row)
1262      VariableBounds(row) = True

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("COpenSolver", "UpdateBound") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Sub BuildModelOptions(succeeded As Boolean, Optional SolveRelaxation As Boolean = False, Optional MinimiseUserInteraction As Boolean = False)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          succeeded = False
          
          ' Ignore integrality check if the user is solving the relaxation
          If Not SolveRelaxation Then
              ' If we are in Excel 2010 or newer, and the user has set "Solve Relaxtion", then we report an error
              Dim IngoreIntConstraints As Boolean
1265          IngoreIntConstraints = False
1266          If Val(Application.Version) >= 14 And GetIgnoreIntegerConstraints(book, sheet) Then
                   If MinimiseUserInteraction Then GoTo RaiseError
                   If Not MsgBox("You have the 'Ignore integer constraints' option enabled in Solver, so OpenSolver cannot proceed with solving. Would you like OpenSolver to disable this option and proceed?", vbCritical + vbYesNo + vbDefaultButton1, "OpenSolver Warning") = vbNo Then
                       SetIgnoreIntegerConstraints False, book, sheet
                   Else
RaiseError:
                       ' Set UserCancelledError because the user has already confirmed the abort if not in silent mode
                       Err.Raise Number:=OpenSolver_UserCancelledError, Description:="You have the Solver option 'ignore integer constraints' turned on. Please turn this off in the Solver (not OpenSolver) options menu and then solve with OpenSolver again or choose OpenSolver's solve relaxed model option"
                   End If
1271          End If
          End If

1272      succeeded = True

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("COpenSolver", "BuildModelOptions") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Sub BuildConstraintMatrix(objIsMissing As Boolean, SparseA() As CIndexedCoeffs, succeeded As Boolean)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

1275      succeeded = False
          
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

1344      succeeded = True
          
ExitSub:
          Application.StatusBar = False
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("COpenSolver", "BuildConstraintMatrix") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Sub ViewSparseA()
                Dim i As Long, j As Long, strA As String
1367            For i = 1 To UBound(SparseA)
1368                For j = 1 To numVars
1369                    strA = strA & CStr(SparseA(i).Coefficient(j)) & "    "
1370                Next j
1371                strA = strA & vbNewLine
1372            Next i
1373            MsgBox (strA)
End Sub

Function SolveModel(SolveRelaxation As Boolean, ShouldMinimiseUserInteraction As Boolean) As Long

          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          ' Track whether to show messages
1375      MinimiseUserInteraction = ShouldMinimiseUserInteraction

          'Check that solver is available
1376      If Not SolverAvailable(Solver, ExternalSolverPathName) Then
1377          Err.Raise Number:=OpenSolver_SolveError, Description:="The specified solver, " & Solver & " could not be found. Make sure it is correctly installed or try another solver."
1378      End If

          'Delete any existing log file
          Dim LogPath As String
1379      LogPath = GetTempFilePath("log1.tmp")
1380      DeleteFileAndVerify LogPath

1381      ModelFilePathName = ModelFilePath(Solver)
          
1382      DeleteFileAndVerify ModelFilePathName

          'Clean up solver specific files
1383      CleanFiles

1384      If Solver = "NOMAD" Then
1385          SolveRelaxedModel = SolveRelaxation
1386          SolveModel = SolveModel_Nomad(SolveRelaxation, Me)
1387      Else
1388          SolveModel = SolveModel_Linear(SolveRelaxation)
1389      End If

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("COpenSolver", "SolveModel") Then Resume
          RaiseError = True
          GoTo ExitFunction

End Function

Function SolveModel_Linear(SolveRelaxation As Boolean) As Long
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

1407      SolveStatus = OpenSolverResult.Unsolved
1408      SolveStatusString = "Unsolved"
1409      SolveStatusComment = ""
1412      SolutionWasLoaded = False
          
          Dim oldCalculationMode As Long
1413      oldCalculationMode = Application.Calculation
1414      Application.Calculation = xlCalculationManual

1417      If ModelStatus <> ModelStatus_Built Then
1418          Err.Raise Number:=OpenSolver_SolveError, Description:="The model cannot be solved as it has not yet been built."
1419      End If

          ' Trap Escape key
1420      Application.EnableCancelKey = xlErrorHandler
          
1421      UpdateStatusBar "OpenSolver: Writing Model to disk... " & numVars & " vars, " & NumRows & " rows.", True
          
          ' Check if we need to request duals from the solver
          Set rConstraintList = GetDuals()
1422      DualsOnSameSheet = (Not rConstraintList Is Nothing)
1424      DualsOnNewSheet = GetDualsNewSheet()
1430      bGetDuals = (IntegerCellsRange Is Nothing And BinaryCellsRange Is Nothing) And _
                      (DualsOnNewSheet Or DualsOnSameSheet) And SolverHasSensitivityAnalysis(Solver)
          
          ' Set up arrays to hold solution values (to avoid dynamically resizing them later)
1434      ReDim pFinalVarValue(numVars)
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
          
          'Delegate to Neos or Local solver
1445      If RunsOnNeos(Solver) Then
1446          SolveModel_LinearNeos ModelFilePathName, SolveRelaxation
1447      Else
1448          SolveModel_LinearLocal ModelFilePathName, SolveRelaxation
1449      End If
          
          If SolutionWasLoaded Then
              AdjustableCells.Value2 = 0
              Dim i As Long
              For i = 1 To numVars
                  AdjustableCells.Worksheet.Range(VarCell(i)).Value2 = FinalVarValue(i)
              Next i
          End If
          
          ' Perform a linearity check unless the user has requested otherwise
          If GetLinearityCheck() Then
              Dim fullLinearityCheckWasPerformed As Boolean
1456          QuickLinearityCheck fullLinearityCheckWasPerformed
1457          If fullLinearityCheckWasPerformed Then
1458              SolveStatus = OpenSolverResult.AbortedThruUserAction
1459              SolveStatusString = "No Solution Found "
1461          End If
1462      End If
            
          '=====================Write Sensitivity Tables=========================================================
1463      If SolverHasSensitivityAnalysis(Solver) Then
1464          If bGetDuals And SolveStatus = OpenSolverResult.Optimal Then
                  'write the duals on the same sheet if the user has picked this option
1465              If DualsOnSameSheet Then WriteConstraintListToSheet rConstraintList
                  'If the user wants a new sheet with the sensitivity data then call the functions that write this
1466              If DualsOnNewSheet Then
                      Dim nameSheet As String
1467                  nameSheet = MakeNewSheet(RawSheetName & " Sensitivity", sheet)
1468                  sheet.Select
1469                  Call WriteConstraintSensitivityTable(nameSheet, Solver)
                      'If you want to end up on the sensitivity page use this next line
                      'Sheets(nameSheet).Select
1470              End If
1471          ElseIf Not bGetDuals And (DualsOnNewSheet Or DualsOnSameSheet) Then
1472              Err.Raise Number:=OpenSolver_SolveError, Description:="Could not get sensitivity analysis due to binary and/or integer constraints." & vbCrLf & vbCrLf & "Turn off sensitivity in the model dialogue or reformulate your model without these constraints." & vbCrLf & vbCrLf & "The " & Solver & " solution has been returned to the sheet." & vbCrLf
1473          End If
1474      End If

ExitFunction:
1481      SolveModel_Linear = SolveStatus    ' Return the main result
1495      Application.Cursor = xlDefault
1496      Application.StatusBar = False ' Resume normal status bar behaviour
1497      Application.ScreenUpdating = True
1498      Application.Calculation = oldCalculationMode
1499      Application.Calculate
1500      Close #1 ' Close any open file; this does not seem to ever give errors
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("COpenSolver", "SolveModel_Linear") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Sub SolveModel_LinearNeos(ModelFilePathName As String, SolveRelaxation As Boolean)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

1503      WriteAMPLFile ModelFilePathName, SolveRelaxation
          If SolveStatus <> OpenSolverResult.Unsolved Then GoTo ExitSub
          
          Dim solution As String
1507      solution = CallNEOS(ModelFilePathName, Solver, MinimiseUserInteraction)

          SolutionWasLoaded = False
          
          Dim openingParen As String, closingParen As String
          ' Extract the solve status
1511      openingParen = InStr(solution, "solve_result")
1512      SolveStatusString = Mid(solution, openingParen + Len("solve_result"))
               
          ' Determine Feasibility
1513      If SolveStatusString Like "*unbounded*" Then
1514          SolveStatus = OpenSolverResult.Unbounded
1515          SolveStatusString = "No Solution Found (Unbounded)"
1517      ElseIf SolveStatusString Like "*infeasible*" Then ' Stopped on iterations or time
1518          SolveStatus = OpenSolverResult.Infeasible
1519          SolveStatusString = "No Feasible Solution"
1522      ElseIf Not SolveStatusString Like "*solved*" Then
1523          openingParen = InStr(solution, ">>>")
1524          If openingParen = 0 Then
1525              openingParen = InStr(solution, "processing commands.")
1526              SolveStatusComment = Mid(solution, openingParen + Len("processing commands."))
1527          Else
1528              closingParen = InStr(solution, "<<<")
1529              SolveStatusComment = Mid(solution, openingParen, closingParen - openingParen)
1530          End If
1531          SolveStatus = OpenSolverResult.ErrorOccurred
1532      Else
1533          SolveStatus = OpenSolverResult.Optimal
1534          SolveStatusString = "Optimal"
              SolutionWasLoaded = True
1536      End If
          
          If SolutionWasLoaded Then
              Dim var As Long, resultLength As Long, start As Long
1538          For var = 1 To numVars
1539              openingParen = InStr(solution, ValidLPFileVarName(VarNames(var)))
1540              closingParen = openingParen + InStr(Mid(solution, openingParen + 1), "_display")
                  start = openingParen + Len(ValidLPFileVarName(VarNames(var))) + 1
1541              resultLength = closingParen - start
                  If resultLength > 0 Then
                      FinalVarValue(var) = ConvertToCurrentLocale(Mid(solution, start, resultLength))
                  Else
                      FinalVarValue(var) = 0
                  End If
                  VarCell(var) = VarNames(var)
1544          Next
          End If

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("COpenSolver", "SolveModel_LinearNeos") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Sub SolveModel_LinearLocal(ModelFilePathName As String, SolveRelaxation As Boolean)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
             
1550      WriteLPFile ModelFilePathName, SolveRelaxation
          If SolveStatus <> OpenSolverResult.Unsolved Then GoTo ExitSub
                  
          ' Get the Solver Options, stored in named ranges with values such as "=0.12"
          ' Because these are NAMEs, they are always in English, not the local language, so get their value using Val
          Dim SolveOptions As SolveOptionsType
1557      GetSolveOptions sheet, SolveOptions
              
1562      UpdateStatusBar "OpenSolver: Solving " & IIf(SolveRelaxation, "Relaxed ", "") & "Model... " & numVars & " vars, " & _
                          NumIntVars & " int vars " & "(" & NumBinVars & " 0/1), " & NumRows & " rows, " & _
                          SolveOptions.MaxTime & "s time limit, " & SolveOptions.Tolerance * 100 & "% tolerance.", True
              
          ' We allow user to define extra parameters
          Dim ExtraParameters As New Dictionary
1563      GetExtraParameters Solver, sheet, ExtraParameters
              
          Dim SolutionFilePathName As String
1568      SolutionFilePathName = SolutionFilePath(Solver)
         
          Dim logCommand As String, logFileName As String
1570      logFileName = "log1.tmp"
1571      logCommand = GetTempFilePath(logFileName)
         
1569      ExternalSolverPathName = CreateSolveScript(Solver, SolutionFilePathName, ExtraParameters, SolveOptions, Me)

              
          Dim ExecutionCompleted As Boolean
          ExternalSolverPathName = MakePathSafe(ExternalSolverPathName) ' fix error for Gurobi with spaces in path name by quoting path
          
          Dim exeResult As Long
1573      ExecutionCompleted = RunExternalCommand(ExternalSolverPathName, logCommand, IIf(SolveOptions.ShowIterationResults And Not MinimiseUserInteraction, SW_SHOWNORMAL, SW_HIDE), True, exeResult) ' Run solver, waiting for completion
1578      If exeResult <> 0 Then
1580          Err.Raise Number:=OpenSolver_SolveError, Description:="The " & Solver & " solver did not complete, but aborted with the error code " & exeResult & "." & vbCrLf & vbCrLf & "The last log file can be viewed under the OpenSolver menu and may give you more information on what caused this error."
1581      End If
              
          SolutionFilePathName = SolutionFilePath(Solver)
1584      If Not FileOrDirExists(SolutionFilePathName) Then
1586          Err.Raise Number:=OpenSolver_SolveError, Description:="The solver did not create a solution file. No new solution is available."
1588      End If
              
          Dim solutionLoaded As Boolean
1589      solutionLoaded = ReadModel(Solver, SolutionFilePathName, Me)

ExitSub:
          Close #1
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("COpenSolver", "SolveModel_LinearLocal") Then Resume
          RaiseError = True
          GoTo ExitSub
          End Sub

' Output the model to an LP format text file. See http://lpsolve.sourceforge.net/5.5/CPLEX-format.htm
Sub WriteLPFile(ModelFilePathName As String, SolveRelaxation As Boolean)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim i As Long, j As Long, var As Long, row As Long, coeff As Double
          Dim commentStart As String  'Character for starting comments for chosen solver
1613      commentStart = "\"
          
          ' Track which variables we have printed so far - only needed when AssumeNonNegativeVars is true
          ' This is so we can print all unused variables as FREE for sensitivity analysis
          ' When AssumeNonNegativeVars is false, we print all vars as FREE anyway, so this isn't needed
          Dim UsedVariables As Collection, VarName As String
          If AssumeNonNegativeVars Then
              Set UsedVariables = New Collection
          End If
          
1615      Open ModelFilePathName For Output As #1 ' supply path with filename
1616      Print #1, commentStart & " Model solved using the solver '" & Solver & "'"
1617      Print #1, commentStart & " Model for sheet " & left(sheetName, Len(sheetName) - 1) ' remove trailing !
          ' Warning: TimeTakenToBuildModel may be negative if we cross midnight! This is unlikely for short build times
1618      Print #1, commentStart & " It took " & IIf(SecondsTakenToBuildModel > 100, SecondsTakenToBuildModel, TimeTakenToBuildModel) & " seconds to build the model."
1619      Print #1, commentStart & " Model has " & NumConstraints & " Excel constraints giving " & NumRows & " constraint rows and " & numVars & " variables."
1620      If SolveRelaxation And (NumBinVars > 0) Then
1621          Print #1, commentStart & " (Formulation for relaxed problem)"
1622      End If
1623      Print #1, IIf(ObjectiveSense = MaximiseObjective, "MAXIMIZE", "MINIMIZE")   ' We use Minimise for both minimisation, and also for seeking a target (TargetObjective)
1624      Print #1, "Obj:";
1625      If ObjectiveSense = TargetObjective Then
              ' We want the objective to achieve some target value; we have no objective; nothing is output
1626          Print #1, ' A new line meaning a blank objective; add a comment to this effect next
1628          Print #1, commentStart & " We have no objective function as the objective must achieve a given target value"
1629      Else
1631          For var = 1 To numVars
1632              If Abs(CostCoeffs(var)) > EPSILON Then
                      VarName = ValidLPFileVarName(VarNames(var))
                      If AssumeNonNegativeVars Then
                          If Not TestKeyExists(UsedVariables, VarName) Then UsedVariables.Add VarName, VarName
                      End If
                      Print #1, " " & StrEx(CostCoeffs(var)) & " " & VarName;
                  End If
              Next var
              'If Abs(objValue) > EPSILON Then
              '    Print #1, " " & StrEx(objValue);
              'End If
1634          Print #1,   ' New line
1635      End If
1636      Print #1,   ' New line
1637      Print #1, "SUBJECT TO"
          
          ' If we are seeking a specific objective value, we add this as a constraint
1638      If ObjectiveSense = TargetObjective Then
1639          Print #1, commentStart & " The objective must achieve a given target value; this constraint enforces this."
              Dim NonTrivialObjective As Boolean
1640          For var = 1 To numVars
1641              If Abs(CostCoeffs(var)) > EPSILON Then
                      VarName = ValidLPFileVarName(VarNames(var))
                      If AssumeNonNegativeVars Then
                          If Not TestKeyExists(UsedVariables, VarName) Then UsedVariables.Add VarName, VarName
1642                  End If
                      Print #1, " " & StrEx(CostCoeffs(var)) & " " & VarName;
1643                  NonTrivialObjective = True
1644              End If
1645          Next var
1646          Print #1, " = " & StrEx(ObjectiveTargetValue)
1647          If Not NonTrivialObjective And ObjectiveTargetValue <> 0 Then
1648              SolveStatus = OpenSolverResult.Infeasible
1649              SolveStatusString = "Infeasible Objective Target"
1650              SolveStatusComment = "The model's objective cell does not depend on the decision variables" _
                        & " and so cannot be adjusted to achieve the target value" & ObjectiveTargetValue & "."
1651              GoTo ExitSub
1652          End If
1653      End If
          
          Dim constraint As Long
          Dim instance As Long
1654      constraint = 1
          
1655      For row = 1 To NumRows
1656          GetConstraintFromRow row, constraint, instance  ' Which Excel constraint are we in, and which instance?
1657          If instance = 1 Then
                  ' We are outputting the first row of a new Excel constraint block; put a comment in the .lp file
1658              Print #1, commentStart & " " & ConstraintSummary(constraint)
1659          End If
         
1660          With SparseA(row)
1661              For i = 1 To .Count
1662                  var = .Index(i)
1663                  coeff = .Coefficient(i)
                      VarName = ValidLPFileVarName(VarNames(var))
                      If AssumeNonNegativeVars Then
                          If Not TestKeyExists(UsedVariables, VarName) Then UsedVariables.Add VarName, VarName
                      End If
1664                  Print #1, " " & StrEx(coeff) & " " & VarName;
1665              Next i
1666          End With
              
1667          If SparseA(row).Count = 0 Then
                ' We have a constraint that does not vary with the decision variables; check it is satisfied
1668            If CheckConstantConstraintIsSatisfied(row, constraint, instance, i, j) Then
                    ' We output the row as a comment
1669                Print #1, commentStart & " (A row with all zero coeffs)";
1670            Else
1671                GoTo ExitSub
1672            End If
1673          End If
1674          If Relation(row) = RelationGE Then
1675              Print #1, " >= ";
1676          ElseIf Relation(row) = RelationLE Then
1677              Print #1, " <= ";
1678          ElseIf Relation(row) = RelationEQ Then
1679              Print #1, " = ";
1680          End If
1681          Print #1, StrEx(RHS(row))
1682      Next row
1683      Print #1,   ' New line
          Dim c As Range
1684      If SolveRelaxation And Not (BinaryCellsRange Is Nothing) Then
1685          Print #1, commentStart & " (Upper bounds of 1 on the relaxed binary variables)"
1686          For Each c In BinaryCellsRange
                  VarName = ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False))
                  If AssumeNonNegativeVars Then
                      If Not TestKeyExists(UsedVariables, VarName) Then UsedVariables.Add VarName, VarName
                  End If
1687              Print #1, VarName; " <= 1"
1688          Next c
1689          Print #1,   ' New line
1690      End If
          
          ' Output the bounds; this should happen before we output the GENERAL or INTEGER sections (at least for CPLEX .lp files)
          ' See http://lpsolve.sourceforge.net/
          ' The LP file assumes lower bounds on all variables unless we tell it otherwise.
1691      If Not AssumeNonNegativeVars Then
              ' We need to make all variables FREE variables (i.e. no lower bounds), except for the Binary variables
1692          Print #1, "BOUNDS"
1693          Print #1, commentStart & "'Assume Non Negative' is FALSE, so default lower bounds of zero are removed from all non-binary variables."
1695          For Each c In SetDifference(AdjustableCells, BinaryCellsRange)
1696              Print #1, " " & ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False)) & " FREE"
1697          Next c
1705          Print #1,   ' New line
1706      Else
              ' If AssumeNonNegative, then we need to apply lower bounds to any variables without explicit lower bounds.
              
              ' Get all bounded variables
              Dim BoundedVariables As Range
              Set BoundedVariables = Nothing
1707          For Each c In AdjustableCells
                  VarName = ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False))
1708              If TestKeyExists(VarLowerBounds, c.Address) Or Not TestKeyExists(UsedVariables, VarName) Then
                      Set BoundedVariables = ProperUnion(BoundedVariables, c)
                  End If
1711          Next c

              ' We need to mark variables with explicit lower bounds as FREE variables (allowing the possibly negative lower bound to be effective).
              ' However, we don't make Binary variables free
1712          If Not BoundedVariables Is Nothing Then
1713              Print #1, "BOUNDS"
1714              Print #1, commentStart & "'Assume Non Negative' is TRUE, so default lower bounds of zero are removed only from non-binary variables already given explicit lower bounds."
1716              For Each c In SetDifference(BoundedVariables, BinaryCellsRange)
1717                  Print #1, " " & ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False)) & " FREE"
1718              Next c
1726              Print #1,   ' New line
1727          End If
1728      End If
          
          ' Output any integer variables
1729      If Not SolveRelaxation And Not (IntegerCellsRange Is Nothing) Then
1750          Print #1, "GENERAL"
1751          For Each c In IntegerCellsRange
1752              Print #1, " " & ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False));
1753          Next c
1754      End If
1755      Print #1, ' New line

          ' Output binary variables
1757      If Not SolveRelaxation And Not (BinaryCellsRange Is Nothing) Then
1758          Print #1, "BINARY"
1759          For Each c In BinaryCellsRange
1760              Print #1, " " & ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False));
1761          Next c
1762          Print #1,   ' New line
1763      End If
1764      Print #1,   ' New line
          
1765      Print #1, "END"
          
ExitSub:
          Close #1
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("COpenSolver", "WriteLPFile") Then Resume
          RaiseError = True
          GoTo ExitSub

End Sub

Sub WriteAMPLFile(ModelFilePathName As String, SolveRelaxation As Boolean)
           Dim RaiseError As Boolean
           RaiseError = False
           On Error GoTo ErrorHandler

           Dim i As Long, j As Long, var As Long, row As Long, coeff As Double, c As Range, Line As String
           Dim VarDic As Collection
1773       Set VarDic = New Collection

           Dim commentStart As String  'Character for starting comments for chosen solver
1775       commentStart = "#"

1777       Open ModelFilePathName For Output As #1 ' supply path with filename
                
           ' Model File - Replace with Data File
1778       Print #1, "# Define our sets, parameters and variables (with names matching those"
1779       Print #1, "# used in defining the data items)"
           
           ' Variables
1780       For var = 1 To numVars
1781           VarDic.Add "", ValidLPFileVarName(VarNames(var))
1782       Next var
           
           ' Sets - Vars
1783       If Not SolveRelaxation Then
1784           If Not IntegerCellsRange Is Nothing Then
1785               For Each c In IntegerCellsRange
1786                   VarDic.Remove (ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False)))
1787                   VarDic.Add ", integer", ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False))
1788               Next c
1789           End If
           End If
           If Not BinaryCellsRange Is Nothing Then
1791           For Each c In BinaryCellsRange
                   VarDic.Remove (ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False)))
1792               If Not SolveRelaxation Then
1793                   VarDic.Add ", binary", ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False))
                   Else
                       VarDic.Add ", <= 1, >= 0", ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False))
                   End If
1794           Next c
1796       End If
           
           Dim constraint As Long
           Dim instance As Long
1797       constraint = 1
           
           ' Reindex bounded variables by relative address so that VarNames(var) can search the collection
           Dim BoundedVariables As New Collection
1798       For Each c In AdjustableCells
1799           If TestKeyExists(VarLowerBounds, c.Address) Then
1800               BoundedVariables.Add c, c.Address(RowAbsolute:=False, ColumnAbsolute:=False)
1801           End If
1802       Next c
           
           ' Intialise Variables
1803       For var = 1 To numVars
1804           Line = "var " & ValidLPFileVarName(VarNames(var)) & VarDic.Item(ValidLPFileVarName(VarNames(var)))
1805           If AssumeNonNegativeVars Then
                   ' If no lower bound has been applied then we need to add >= 0
1806               If Not TestKeyExists(BoundedVariables, VarNames(var)) Then
1807                   Line = Line & " >= 0"
1808               End If
1809           End If
1810           Print #1, Line & ";"
1811       Next var
           
           ' Parameter - Costs
1812       Print #1,   ' New line
1813       Line = "  "
1814       For var = 1 To numVars
               If Abs(CostCoeffs(var)) > EPSILON Then
1815               Line = Line & ValidLPFileVarName(VarNames(var)) & "*" & StrEx(CostCoeffs(var))
1816               If var < numVars Then
1817                   Line = Line & " + "
                   End If
1818           End If
1819       Next var
           Line = Line & " " & StrEx(objValue)
           
           ' Objective function replaced with constraint if
1820       If ObjectiveSense = TargetObjective Then
1821           Print #1, commentStart & " We have no objective function as the objective must achieve a given target value"
1822           Print #1,
               
1823           Print #1, commentStart & " The objective must achieve a given target value; this constraint enforces this."
1824           Print #1, "subject to TargetConstr:"
1825           Print #1, Line & " = " & StrEx(ObjectiveTargetValue) & ";"
1826       Else
               ' Determine objective direction
1827           If ObjectiveSense = MaximiseObjective Then
1828              Print #1, "maximize Total_Cost:"
1829           Else
1830              Print #1, "minimize Total_Cost:"
1831           End If
               
1832           Print #1, Line & ";"
1833           Print #1,   ' New line
1834       End If
           
           ' Subject to Constraints
1835       For row = 1 To NumRows
1836           Line = "   "
1837           GetConstraintFromRow row, constraint, instance  ' Which Excel constraint are we in, and which instance?
1838           If instance = 1 Then
                   ' We are outputting the first row of a new Excel constraint block; put a comment in the .lp file
1839               Print #1, commentStart & " " & ConstraintSummary(constraint)
1840           End If
               
               ' Gather variables
1841           With SparseA(row)
1842               For i = 1 To .Count
1843                   var = .Index(i)
1844                   coeff = .Coefficient(i)
1845                   Line = Line & StrEx(coeff) & " * " & ValidLPFileVarName(VarNames(var)) & " "
1846               Next i
1847           End With
              
              ' Check sense
1848          If Relation(row) = RelationGE Then
1849              Line = Line & ">= "
1850          ElseIf Relation(row) = RelationLE Then
1851              Line = Line & "<= "
1852          ElseIf Relation(row) = RelationEQ Then
1853              Line = Line & "== "
1854          End If
              
              ' Add RHS
1855          Line = Line & StrEx(RHS(row))
              
1856          If SparseA(row).Count = 0 Then
                  ' We have a constraint that does not vary with the decision variables; check it is satisfied
1857              If CheckConstantConstraintIsSatisfied(row, constraint, instance, i, j) Then
                      'We output the row as a comment
1858                  Print #1, commentStart & " (A row with all zero coeffs)" & Line & ";"
1859              Else
1860                  GoTo ExitSub
1861              End If
1862          Else
                  'Output the constraint
1863              Print #1, "subject to c" & row & ":"
1864              Print #1, Line
1865              Print #1, ";"
1866          End If
             
1867          Print #1,   ' New line
1868       Next row
           
           ' Run Commands
1869       Print #1,   ' New line
1870       Print #1, commentStart & " Solve the problem"
1871       Print #1, "option solver cbc;"
1872       Print #1, "solve;"
1873       Print #1,   ' New line
           ' Display variables
1874       For var = 1 To numVars
1875           Print #1, "_display " & ValidLPFileVarName(VarNames(var)) & ";"
1876       Next var
1877       Print #1, "_display solve_result;"
1878       Print #1,   ' New line
           
ExitSub:
           Close #1
           If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
           Exit Sub

ErrorHandler:
           If Not ReportError("COpenSolver", "WriteAMPLFile") Then Resume
           RaiseError = True
           GoTo ExitSub

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
1898          SolveStatusComment = "The model contains an Excel constraint '" & ConstraintSummary(constraint) & "' for which instance " _
                     & instance & " does not depend on the decision variables and is not satisfied." & vbCrLf _
                    & "Constraint specifies: " _
                     & "LHS=" & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "=" & LHSValue _
                    & " " & SolverRelationAsString(Relation(row)) & " " _
                    & "RHS=" & RHSstring & "=" & RHSValue
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
          ' We have already raised non-linearity to the user, don't show this again
1904      If SolveStatus <> OpenSolverResult.Optimal And SolveStatus <> OpenSolver.NotLinear Then
1905          If SolveStatus = OpenSolverResult.AbortedThruUserAction Then
                  ' Do nothing; user has already interacted with some dialogs
1906          Else
1909              MsgBox "OpenSolver could not find an optimal solution, and reported:" & vbCrLf & _
                         SolveStatusString & vbCrLf & vbCrLf & _
                         IIf(SolveStatusComment = "", "", SolveStatusComment & vbCrLf & vbCrLf) & _
                         IIf(SolutionWasLoaded, "The solution generated has been loaded into the spreadsheet.", "No solution was available to load into the spreadsheet.")
1910          End If
1911      End If
End Sub

Sub WriteConstraintListToSheet(r As Range)
          ' Write a list of all the constraints in a column at cell r
          ' TODO: This will not correctly handle constraints on another sheet
          
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

1916      r.Cells(1, 1).Value2 = "Cons"
1917      r.Cells(1, 2).Value2 = "SP"
1918      r.Cells(1, 3).Value2 = "Inc"
1919      r.Cells(1, 4).Value2 = "Dec"
          
          Dim constraint As Long, row As Long, instance As Long
1920      constraint = 1
1921      For row = 1 To NumRows
1922          GetConstraintFromRow row, constraint, instance  ' Which Excel constraint are we in, and which instance?
              
              Dim UnusedConstraint As Boolean
1923          UnusedConstraint = SparseA(row).Count = 0
              
              Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
1924          GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
1925          If Not RHSCellRange Is Nothing Then
1926              RHSstring = RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
1927          Else
1928              RHSstring = ConvertToCurrentLocale(StripWorksheetNameAndDollars(RHSstring, ActiveSheet))
1930          End If
              
              Dim s As String
1931          s = IIf(UnusedConstraint, "", "") & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & _
                  SolverRelationAsString(Relation(row)) & RHSstring & IIf(UnusedConstraint, "", "")
          
1932          r.Cells(row + 1, 1).value = s
1933          r.Cells(row + 1, 2).Value2 = ZeroIfSmall(ShadowPrice(row))
1934          r.Cells(row + 1, 3).Value2 = ZeroIfSmall(IncreaseCon(row))
1935          r.Cells(row + 1, 4).Value2 = ZeroIfSmall(DecreaseCon(row))
1936      Next row

          'Write the variable duals
1937      row = row + 2
1938      r.Cells(row, 1).Value2 = "Vars"
1939      r.Cells(row, 2).Value2 = "RC"
1940      r.Cells(row, 3).Value2 = "Inc"
1941      r.Cells(row, 4).Value2 = "Dec"
1942      row = row + 1
          Dim i As Long
1943      For i = 1 To UBound(pReducedCosts)
1944          r.Cells(row, 1).Value2 = VarCell(i)
1945          r.Cells(row, 2).Value2 = ZeroIfSmall(ReducedCosts(i))
1946          r.Cells(row, 3).Value2 = ZeroIfSmall(IncreaseVar(i))
1947          r.Cells(row, 4).Value2 = ZeroIfSmall(DecreaseVar(i))
1948          row = row + 1
1949      Next i

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("COpenSolver", "WriteConstraintListToSheet") Then Resume
          RaiseError = True
          GoTo ExitSub

End Sub

Sub WriteModelToSheet()
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim oldCalculationMode As Long
1950      oldCalculationMode = Application.Calculation
1951      Application.Calculation = xlCalculationManual
          
1952      UpdateStatusBar "OpenSolver: Writing Model to worksheet 'Model'... " & numVars & " vars, " & NumRows & " rows.", True

          Dim var As Long, i As Long, row As Long, coeff As Double
          
1953      For var = 1 To numVars
1954          Worksheets("Model").Cells(1, 1 + var) = VarNames(var)
1955      Next var
          
          Dim c As Range
1956      var = 0
          
1957      For Each c In AdjustableCells
1958              var = var + 1
1959              If Not Intersect(c, IntegerCellsRange) Is Nothing Then
1960                  Worksheets("Model").Cells(2, 1 + var) = "Int"
1961              ElseIf Not Intersect(c, BinaryCellsRange) Is Nothing Then
1962                  Worksheets("Model").Cells(2, 1 + var) = "Bin"
1963              End If
1964      Next c
           
      '     For var = 1 To NumVars
      '        If VarTypes(i) = RelationInt Then
      '            Worksheets("Model").Cells(2, 1 + var) = "Int"
      '        ElseIf VarTypes(i) = RelationBin Then
      '            Worksheets("Model").Cells(2, 1 + var) = "Bin"
      '        End If
      '    Next var
      '

1965      For var = 1 To numVars
1966          Worksheets("Model").Cells(3, 1 + var) = CostCoeffs(var)
1967      Next var
              'For var = 1 To NumVars
              '    If a(row, var) <> 0 Then
              '        Worksheets("Model").Cells(4 + row, 1 + var) = a(row, var)
              '    End If
              'Next var
1968      With SparseA(row)
1969          For i = 1 To .Count
1970              var = .Index(i)
1971              coeff = .Coefficient(i)
1972              Worksheets("Model").Cells(4 + row, 1 + var) = coeff
1973          Next i
1974      End With
1975      For row = 1 To NumRows
1976          Worksheets("Model").Cells(4 + row, 1 + numVars + 1) = SolverRelationAsUnicodeChar(Relation(row))
1977          Worksheets("Model").Cells(4 + row, 1 + numVars + 2) = RHS(row)
1978      Next row
          
ExitSub:
1979      Application.StatusBar = False
1980      Application.ScreenUpdating = True
1981      Application.Calculation = oldCalculationMode
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("COpenSolver", "WriteModelToSheet") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Sub ClearQuickSolve()
1992      Set QuickSolve = Nothing
End Sub

Function InitializeQuickSolve() As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

1993      InitializeQuickSolve = False
1994      If ModelStatus <> ModelStatus_Built Then
1995          MsgBox "QuickSolve Initialisation Failed: The model cannot be solved as it has not yet been built.", , "OpenSolver" & sOpenSolverVersion & " Error"
1996          GoTo ExitFunction
1997      End If

1998      Set QuickSolve = New CQuickSolve
1999      If QuickSolve.AnalyseParameters(Me, sheet, NumConstraints, RHS, LHSRange, RHSRange, RHSFormula, RHSType, LHSOriginalValues, RHSOriginalValues) Then
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

Function DoQuickSolve(Optional MinimiseUserInteraction As Boolean = False) As Long
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          ' Update a previously built model
2012      DoQuickSolve = OpenSolverResult.Unsolved
2013      If QuickSolve Is Nothing Then
2014          MsgBox "Error: The quick solve data has not been initialized. Please choose Intialize Quick Solve, and then try again.", , "OpenSolver" & sOpenSolverVersion & " Error"
2015      Else
2016          If QuickSolve.UpdateModelRHS(sheet, RHS) Then
2017              DoQuickSolve = SolveModel(False, MinimiseUserInteraction)
2018          End If
2019      End If

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("COpenSolver", "DoQuickSolve") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

' Highlight all constraints (and the objective) that are non-linear using our standard model highlighting, but showing only individual cells, not ranges
Sub HighlightNonLinearities(RowIsNonLinear() As Boolean, ObjectiveIsNonLinear As Boolean)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim constraint As Long, row As Long, instance As Long
2028      If SheetHasOpenSolverHighlighting(ActiveSheet) Then
2029          HideSolverModel
2030      End If
2031      DeleteOpenSolverShapes ActiveSheet
2032      InitialiseHighlighting
2033      constraint = 1
2034      For row = 1 To NumRows
2035          If RowIsNonLinear(row) Then
                  Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
2036              GetConstraintFromRow row, constraint, instance
2037              GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
2038              RHSstring = StripWorksheetNameAndDollars(RHSstring, LHSCellRange.Worksheet) ' Strip any worksheet name and $'s from the RHS (useful if it is a formula)
                  Dim RHSisRange As Boolean
2039              RHSisRange = RHSType(constraint) = MultiCellRange Or RHSType(constraint) = SingleCellRange
2040              HighlightConstraint LHSCellRange.Worksheet, LHSCellRange, RHSCellRange, RHSstring, Relation(row), 0  ' Show either a value or a formula
2041          End If
2042      Next row
2043      If ObjectiveIsNonLinear Then
              Dim ObjType As ObjectiveSenseType, ObjectiveTargetValue As Double
2044          ObjType = GetObjectiveSense(book, sheet)
2046          If ObjType = TargetObjective Then ObjectiveTargetValue = GetObjectiveTargetValue(book, sheet)
2047          AddObjectiveHighlighting ObjRange, ObjType, ObjectiveTargetValue
2048      End If

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("COpenSolver", "HighlightNonLinearities") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Sub CheckLinearityOfModel()
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim row As Long, i As Long, RowIsNonLinear() As Boolean
          Dim ValueZero() As CIndexedCoeffs, ValueOne() As CIndexedCoeffs, ValueTen() As CIndexedCoeffs, OriginalSolutionValues() As Variant
          Dim NonLinearInformation As String
          Dim ObjectiveCoeffsZero() As Double, ObjectiveCoeffsOne() As Double, ObjectiveCoeffsTen() As Double
          Dim ObjectiveFunctionConstantZero As Double, ObjectiveFunctionConstantOne As Double, ObjectiveFunctionConstantTen As Double
          
2049      ReDim SolutionValues(numVars) As Double
2050      If NumRows > 0 Then ReDim Preserve ValueZero(NumRows) As CIndexedCoeffs
          
2051      NonLinearInformation = ""
          
          ' Remember the original decision variable values (in a variant array to handle multiple areas)
2052      OriginalSolutionValues = GetRangeValues(AdjustableCells)
          
2053      ReDim ObjectiveCoeffsZero(numVars) As Double, ObjectiveCoeffsOne(numVars) As Double, ObjectiveCoeffsTen(numVars) As Double
2054      If NumRows > 0 Then ReDim RowIsNonLinear(NumRows) As Boolean
          'Build each matrix where the decision variables start at zero (ValueZero()), one (ValueOne()) and ten (ValueTen())
2055      For row = 1 To NumRows
2056          Set ValueZero(row) = SparseA(row).Clone
2057      Next row
2058      For i = 1 To numVars
2059          ObjectiveCoeffsZero(i) = CostCoeffs(i)
2060      Next i
2061      ObjectiveFunctionConstantZero = ObjectiveFunctionConstant
          
2062      BuildModelFromSolverData 1
2063      If NumRows > 0 Then ReDim Preserve ValueOne(NumRows) As CIndexedCoeffs
2064      For row = 1 To NumRows
2065          Set ValueOne(row) = SparseA(row).Clone
2066      Next row
2067      For i = 1 To numVars
2068          ObjectiveCoeffsOne(i) = CostCoeffs(i)
2069      Next i
2070      ObjectiveFunctionConstantOne = ObjectiveFunctionConstant
          
2071      BuildModelFromSolverData 10
2072      If NumRows > 0 Then ReDim Preserve ValueTen(NumRows) As CIndexedCoeffs
2073      For row = 1 To NumRows
2074          Set ValueTen(row) = SparseA(row).Clone
2075      Next row
2076      For i = 1 To numVars
2077          ObjectiveCoeffsTen(i) = CostCoeffs(i)
2078      Next i
2079      ObjectiveFunctionConstantTen = ObjectiveFunctionConstant
          
          Dim constraint As Long, instance As Long
2080      constraint = 1
          
          'TODO: These tests should not have an AND, and is the model build code valid if we just shift the zero point?
          
          Dim NumEntries As Long, ValueZeroCounter As Long, ValueOneCounter As Long, ValueTenCounter As Long
          Dim FirstVar As Boolean, NonLinearityCount As Long
          'Go through each row and check each coefficient individually. if it is not within the tolerance the its nonlinear
2081      For row = 1 To NumRows
2082          RowIsNonLinear(row) = False
              
              'This is used to display the constriant
2083          FirstVar = True
2084          ValueZeroCounter = ValueZero(row).Count
2085          ValueOneCounter = ValueOne(row).Count
2086          ValueTenCounter = ValueTen(row).Count
              'find out how many variables it dependent on
2087          NumEntries = ChooseMaximum(ValueZeroCounter, ValueOneCounter, ValueTenCounter)
2088          For i = 1 To NumEntries
2089              If TestExistanceOfEntry(ValueZero(row), ValueOne(row), ValueTen(row), i) Then
                      'do a ratio test
2090                  If Abs(ValueOne(row).Coefficient(i) - ValueZero(row).Coefficient(i)) / (1 + Abs(ValueOne(row).Coefficient(i))) > EPSILON _
                      And Abs(ValueOne(row).Coefficient(i) - ValueTen(row).Coefficient(i)) / (1 + Abs(ValueOne(row).Coefficient(i))) > EPSILON Then
2091                      GetConstraintFromRow row, constraint, instance
2092                      If NonLinearityCount <= 10 Then AddNonLinearInfoToString ValueOne(row).Index(i), NonLinearInformation, FirstVar, constraint, instance
2093                      FirstVar = False
2094                      RowIsNonLinear(row) = True
2095                  End If
2096              Else
2097                  GetConstraintFromRow row, constraint, instance
                      Dim VariableIndex As Long
2098                  VariableIndex = GetEntry(ValueZero(row), ValueOne(row), ValueTen(row), i)
2099                  If NonLinearityCount <= 10 Then AddNonLinearInfoToString VariableIndex, NonLinearInformation, FirstVar, constraint, instance
2100                  FirstVar = False
2101                  RowIsNonLinear(row) = True
2102              End If
2103          Next i
2104          If RowIsNonLinear(row) Then NonLinearityCount = NonLinearityCount + 1
2105      Next row
          
2106      If NonLinearityCount > 10 Then
2107          NonLinearInformation = NonLinearInformation & vbNewLine & " and " & CStr(NonLinearityCount - 10) & " other instances."
2108      End If

          Dim ObjectiveIsNonLinear As Boolean
2109      ObjectiveIsNonLinear = False
2110      For i = 1 To numVars
2111          If Abs(ObjectiveCoeffsZero(i) - ObjectiveCoeffsOne(i)) / (1 + Abs(ObjectiveCoeffsZero(i))) > EPSILON _
              And Abs(ObjectiveCoeffsOne(i) - ObjectiveCoeffsTen(i)) / (1 + Abs(ObjectiveCoeffsOne(i))) > EPSILON Then
2112              If Not ObjectiveIsNonLinear Then
2113                  ObjectiveIsNonLinear = True
2114                  NonLinearInformation = NonLinearInformation & vbNewLine & vbNewLine & "The objective function is nonlinear in the following variables: " & VarNames(i)
2115              Else
2116                  NonLinearInformation = NonLinearInformation & " , " & VarNames(i)
2117              End If
                   
2118          End If
2119      Next i
          
          'Put the solution back on the sheet
2120      SetRangeValues AdjustableCells, OriginalSolutionValues

          NonLinearInformation = TrimBlankLines(NonLinearInformation)
          If NonLinearInformation = "" Then
              NonLinearInformation = "There have been no instances of nonlinearity found in this model. Some models can generate warnings of non-linearity " & _
                                     "because of numerical errors that accumulate in the spreadsheet. OpenSolver's non-linearity check can be disabled under OpenSolver's " & _
                                     "Options settings."
          End If
          
          'display dialog to user
          frmNonlinear.SetLinearityResult NonLinearInformation, False
2139      frmNonlinear.Show
          
2140      If frmNonlinear.chkHighlight.value = True Then
2141          HighlightNonLinearities RowIsNonLinear, ObjectiveIsNonLinear
2142      End If

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("COpenSolver", "CheckLinearityOfModel") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Function TestExistanceOfEntry(ValueZero, ValueOne, ValueTen, i) As Boolean
          'Check if ALL the indices exist
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

2144      If ValueZero.Index(i) <> 0 And ValueOne.Index(i) <> 0 And ValueTen.Index(i) <> 0 Then
2145          TestExistanceOfEntry = True
2146      End If

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
2147      If Err.Number = 9 Then
2148          TestExistanceOfEntry = False
              Resume ExitFunction
2149      End If

          If Not ReportError("COpenSolver", "TestExistanceOfEntry") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

'Return the i'th entry from any one of these arrays; it may not exist in all of them
Function GetEntry(ValueZero, ValueOne, ValueTen, i) As Long
2150      On Error Resume Next
2151      If i <= ValueZero.Count Then
2152          GetEntry = ValueZero.Index(i)
2153      ElseIf i <= ValueOne.Count Then
2154          GetEntry = ValueOne.Index(i)
2155      ElseIf i <= ValueTen.Count Then
2156          GetEntry = ValueTen.Index(i)
2157      End If
End Function

Sub AddNonLinearInfoToString(var As Long, NonLinearInformation As String, FirstVar As Boolean, constraint As Long, instance As Long)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

2158      If FirstVar = True Then
2159          If LHSType(constraint) = SolverInputType.SingleCellRange Then
2160              NonLinearInformation = NonLinearInformation & vbNewLine & "In the constraint: " & ConstraintSummary(constraint) & "," & vbNewLine & "  the model appears to be non-linear in the decision variables: " & VarNames(var)
2161          Else
                  NonLinearInformation = NonLinearInformation & vbNewLine & "In instance " & instance & " of the constraint: " & ConstraintSummary(constraint) & "," & vbNewLine & "  the model appears to be non-linear in the following decision variables: " & VarNames(var)
2163          End If
2164      Else
2165          NonLinearInformation = NonLinearInformation & ", " & VarNames(var)
2166      End If

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("COpenSolver", "AddNonLinearInfoToString") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Function ChooseMaximum(EntriesZero As Long, EntriesOne As Long, EntriesTen As Long) As Long
2167      ChooseMaximum = Max_Int(Max_Int(EntriesZero, EntriesOne), EntriesTen)
End Function

Sub QuickLinearityCheck(fullLinearityCheckWasPerformed As Boolean)
' Returns false if a full check was performed by the user, meaning the model result is no longer valid.
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

2179      fullLinearityCheckWasPerformed = False
          
          Dim row As Long ', i As Long
          Dim NonLinearInfo As String, NonlinearConstraint As Boolean
          'Dim x As Double,
          Dim ExpectedValue As Double, SolutionValue As Double, SolutionValueLHS As Double, SolutionValueRHS As Double
          Dim constraint As Long, i As Long, j As Long
          'Dim sLHS As String, sRHS As String
          'Dim LHSArray As Variant
          Dim CurrentLHSValues As Variant, CurrentRHSValues As Variant
          Dim RowIsNonLinear() As Boolean
2180      If NumRows > 0 Then ReDim RowIsNonLinear(NumRows) As Boolean
          
2181      If Not ForceCalculate("Warning: The worksheet calculation did not complete during the linearity test, and so the test may not be correct. Would you like to retry?") Then GoTo ExitSub
          
          ' Get all the decision variable values off the sheet
          Dim DecisionVariableValues() As Double
2184      ReDim DecisionVariableValues(numVars)
2185      DecisionVariableValues = GetDecisionVariableValuesOffSheet
          
2186      NonLinearInfo = ""
          Dim NonLinearityCount As Long
          
2187      row = 1
2188      For constraint = 1 To NumConstraints
2189          If Not LHSRange(constraint) Is Nothing Then ' Skip INT and BINARY constraint
                  ' Get current value(s) for LHS and RHS of this constraint off the sheet
2190              GetCurrentConstraintValues constraint, CurrentLHSValues, CurrentRHSValues
                  
2191              If RHSType(constraint) <> SolverInputType.MultiCellRange Then
2192                  SolutionValueRHS = CurrentRHSValues
2193              End If
                  
                  Dim instance As Long
2194              instance = 0
2195              For i = 1 To UBound(CurrentLHSValues, 1)
2196                  For j = 1 To UBound(CurrentLHSValues, 2)
2197                      instance = instance + 1
2198                      SolutionValueLHS = CurrentLHSValues(i, j)
2199                      If RHSType(constraint) = SolverInputType.MultiCellRange Then
                              '---------------------------------------------------------------
                              'Check whether the LHS and RHS are parallel or perpendicular
                              '---------------------------------------------------------------
2200                          If UBound(CurrentLHSValues, 1) = UBound(CurrentRHSValues, 1) Then
2201                              SolutionValueRHS = CurrentRHSValues(i, j)
2202                          Else
2203                              SolutionValueRHS = CurrentRHSValues(j, i)
2204                          End If
2205                      End If
2206                      SolutionValue = SolutionValueLHS - SolutionValueRHS
                      
                          'Find out what we expect the value to be from Ax = b. We track the maximum value we encounter during the calculation
                          'so that we have some idea of the errors we might expect
                          Dim maxValueInCalculation As Double
2207                      maxValueInCalculation = 0
2208                      ExpectedValue = SparseA(row).Evaluate_RecordPrecision(DecisionVariableValues, maxValueInCalculation) - RHS(row)
2209                      If Abs(RHS(row)) > maxValueInCalculation Then maxValueInCalculation = Abs(RHS(row))
          
                          ' do a ratio test
2210                      If Abs(ExpectedValue - SolutionValue) / (1 + Abs(ExpectedValue)) > Max_Double(EPSILON, EPSILON * maxValueInCalculation) Then
2211                          If NonLinearInfo = "" Then NonLinearInfo = "The following constraint(s) do not appear to be linear: "
2212                          If NonLinearityCount <= 10 Then
2213                              NonLinearInfo = NonLinearInfo & vbNewLine & ConstraintSummary(constraint)
                              
                                  Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
2214                              GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
                                  ' If the RHS is a range, we show its address; if not, RHSString contains the RHS's constant or formula
2215                              If Not RHSCellRange Is Nothing Then RHSstring = RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
2216                              NonLinearInfo = NonLinearInfo & ": instance " _
                                         & instance _
                                         & ", LHS=" & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) _
                                         & ", RHS=" & RHSstring _
                                         & ", " & ExpectedValue _
                                         & "<>" & SolutionValue
                              
2217                          End If
2218                          NonLinearityCount = NonLinearityCount + 1
2219                          RowIsNonLinear(row) = True
2220                          NonlinearConstraint = True
2221                      End If
                    
2222                      row = row + 1
2223                  Next j
2224              Next i
2225          End If
2226      Next constraint
2227      If NonLinearityCount > 10 Then
2228          NonLinearInfo = NonLinearInfo & vbNewLine & " and " & CStr(NonLinearityCount - 10) & " other constraints."
2229      End If

          'check objective function for linearity
          Dim CalculatedObjValue As Double, ObservedObjValue As Double, ObjectiveIsNonLinear As Boolean
          If ObjRange Is Nothing Then
              ObservedObjValue = 0
          Else
2233          ObservedObjValue = ObjRange.Value2
          End If
          
2236      CalculatedObjValue = CalcObjFnValue(DecisionVariableValues)
2237      ObjectiveIsNonLinear = Abs(CalculatedObjValue - ObservedObjValue) / (1 + Abs(CalculatedObjValue)) > EPSILON
2238      If ObjectiveIsNonLinear Then
2239         NonLinearInfo = "The objective function is not linear." & vbNewLine & vbNewLine & NonLinearInfo
2240      End If
          
          'Set the userform up and display any information on nonlinear constraints
2241      If NonLinearInfo <> "" Then
2242          SolveStatus = NotLinear
2243          If Not MinimiseUserInteraction Then
                  NonLinearInfo = "WARNING : " & vbNewLine & TrimBlankLines(NonLinearInfo)
                  frmNonlinear.SetLinearityResult NonLinearInfo, True
2263              frmNonlinear.Show
              
                  'showing the nonlinear constraints
2264              If frmNonlinear.chkHighlight.value = True Then
2265                  HighlightNonLinearities RowIsNonLinear, ObjectiveIsNonLinear
2266              End If
2267              If frmNonlinear.chkFullCheck.value = True Then
                      'Full linearity check run
2268                  CheckLinearityOfModel
2269                  fullLinearityCheckWasPerformed = True
2270              End If
2271          End If
2272      End If

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("COpenSolver", "QuickLinearityCheck") Then Resume
          RaiseError = True
          GoTo ExitSub
          
End Sub

Sub WriteConstraintSensitivityTable(nameSheet As String, Solver As String)
'Writes out the sensitivity table on a new page (like the solver sensitivity report)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim Column As Long, row As Long, i As Long
          
2273      Sheets(nameSheet).Cells(1, 1) = "OpenSolver Sensitivity Report - " & Solver
2274      Sheets(nameSheet).Cells(2, 1) = "Worksheet: [" & ActiveWorkbook.Name & "] " & ActiveSheet.Name
2275      Sheets(nameSheet).Cells(3, 1) = "Report Created: " & Now()
          
2276      Column = 2
2277      row = 6
          
          Dim sheet As String, headings As Variant
          
2278      sheet = ActiveSheet.Name
2279      headings = Array("Cells", "Name", "Final Value", "Reduced Costs", "Objective Value", "Allowable Increase", "Allowable Decrease")
          'headings for the variable table
2280      Sheets(nameSheet).Cells(row - 1, Column - 1) = "Decision Variables"
2281      For i = 1 To UBound(headings)
2282          Sheets(nameSheet).Cells(row, Column + i - 1) = headings(i)
2283      Next i
2284      row = row + 1
          
          Dim numVars As Double
2285      numVars = UBound(pReducedCosts)
          'put the values into the variable table
2286      For i = 1 To numVars
2287          Sheets(nameSheet).Cells(row, Column) = VarCell(i)
2288          Sheets(nameSheet).Cells(row, Column + 2) = ZeroIfSmall(FinalVarValue(i))
2289          Sheets(nameSheet).Cells(row, Column + 3) = ZeroIfSmall(ReducedCosts(i))
2290          Sheets(nameSheet).Cells(row, Column + 4) = ZeroIfSmall(CostCoeffs(i))
2291          Sheets(nameSheet).Cells(row, Column + 5) = ZeroIfSmall(IncreaseVar(i))
2292          Sheets(nameSheet).Cells(row, Column + 6) = ZeroIfSmall(DecreaseVar(i))
2293          Sheets(nameSheet).Cells(row, Column + 1) = findName(sheet, VarCell(i))
2294          row = row + 1
2295      Next i
          
2296      row = row + 2
          
2297      headings(4) = "Shadow Price"
2298      headings(5) = "RHS Value"

          'Headings for constraint table
2299      Sheets(nameSheet).Cells(row - 1, Column - 1) = "Constraints"
2300      For i = 1 To UBound(headings)
2301          Sheets(nameSheet).Cells(row, Column + i - 1) = headings(i)
2302      Next i

2303      row = row + 1
          Dim constraint As Long, instance As Long
2304      constraint = 1

          'Values for constraint table
2305      For i = 1 To NumRows
2306          Sheets(nameSheet).Cells(row, Column + 2) = ZeroIfSmall(FinalValue(i))
2307          Sheets(nameSheet).Cells(row, Column + 3) = ZeroIfSmall(ShadowPrice(i))
2308          Sheets(nameSheet).Cells(row, Column + 4) = ZeroIfSmall(RHS(i))
2309          Sheets(nameSheet).Cells(row, Column + 5) = ZeroIfSmall(IncreaseCon(i))
2310          Sheets(nameSheet).Cells(row, Column + 6) = ZeroIfSmall(DecreaseCon(i))
              'This finds the range for cells of each constraint (similar to WriteConstraintListToSheet)
2311          GetConstraintFromRow i, constraint, instance  ' Which Excel constraint are we in, and which instance?
              Dim UnusedConstraint As Boolean
2312          UnusedConstraint = SparseA(i).Count = 0
              Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
2313          GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
2314          If Not RHSCellRange Is Nothing Then
2315              RHSstring = RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
2316          Else
2317              RHSstring = ConvertToCurrentLocale(StripWorksheetNameAndDollars(RHSstring, ActiveSheet))
2319          End If
              Dim s As String
2320          s = IIf(UnusedConstraint, "", "") & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & _
              SolverRelationAsString(Relation(i)) & RHSstring & IIf(UnusedConstraint, "", "")
              'Cell Range for each constraint
2321          Sheets(nameSheet).Cells(row, Column).value = s
              'Finds the nearest name for the constraint
2322          Sheets(nameSheet).Cells(row, Column + 1) = findName(sheet, LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False))
2323          row = row + 1
2324      Next i

          'Format the sensitivity table
2325      FormatSensitivityTable nameSheet, row, numVars

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("COpenSolver", "WriteConstraintSensitivityTable") Then Resume
          RaiseError = True
          GoTo ExitSub
          
End Sub

Sub FormatSensitivityTable(nameSheet As String, row As Long, numVars As Double)
'Formats the sensitivity table on the new page with borders and bold writing
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

2326      Application.ScreenUpdating = False
          Dim sheet As String, startRow As String
2327      sheet = ActiveSheet.Name
2328      Sheets(nameSheet).Select
2329      startRow = 6
          
2330      Sheets(nameSheet).Cells.EntireColumn.AutoFit
2331      Columns("A:A").ColumnWidth = 5
2332      With Sheets(nameSheet).Range(Cells(2, 2), Cells(row, 8))
2333          .HorizontalAlignment = xlCenter
2334      End With
          
          'Create the borders for the constraint table
2335      Sheets(nameSheet).Range(Cells(startRow, 2), Cells(startRow + numVars, 8)).Select
2336      With Selection.Borders(xlEdgeLeft)
2337          .LineStyle = xlContinuous
2338          .ColorIndex = 0
2339          .TintAndShade = 0
2340          .Weight = xlMedium
2341      End With
2342      With Selection.Borders(xlEdgeTop)
2343          .LineStyle = xlContinuous
2344          .ColorIndex = 0
2345          .TintAndShade = 0
2346          .Weight = xlMedium
2347      End With
2348      With Selection.Borders(xlEdgeBottom)
2349          .LineStyle = xlContinuous
2350          .ColorIndex = 0
2351          .TintAndShade = 0
2352          .Weight = xlMedium
2353      End With
2354      With Selection.Borders(xlEdgeRight)
2355          .LineStyle = xlContinuous
2356          .ColorIndex = 0
2357          .TintAndShade = 0
2358          .Weight = xlMedium
2359      End With
2360      With Selection.Borders(xlInsideVertical)
2361          .LineStyle = xlContinuous
2362          .Weight = xlThin
2363      End With
          
          'Create the borders for the variable table
2364      Sheets(nameSheet).Range(Cells(numVars + startRow + 3, 2), Cells(row - 1, 8)).Select
2365      With Selection.Borders(xlEdgeLeft)
2366          .LineStyle = xlContinuous
2367          .ColorIndex = 0
2368          .TintAndShade = 0
2369          .Weight = xlMedium
2370      End With
2371      With Selection.Borders(xlEdgeTop)
2372          .LineStyle = xlContinuous
2373          .ColorIndex = 0
2374          .TintAndShade = 0
2375          .Weight = xlMedium
2376      End With
2377      With Selection.Borders(xlEdgeBottom)
2378          .LineStyle = xlContinuous
2379          .ColorIndex = 0
2380          .TintAndShade = 0
2381          .Weight = xlMedium
2382      End With
2383      With Selection.Borders(xlEdgeRight)
2384          .LineStyle = xlContinuous
2385          .ColorIndex = 0
2386          .TintAndShade = 0
2387          .Weight = xlMedium
2388      End With
2389      With Selection.Borders(xlInsideVertical)
2390          .LineStyle = xlContinuous
2391          .Weight = xlThin
2392      End With
          
          'Bold the constraint table headings and make them blue as well as put a border around them
2393      Sheets(nameSheet).Range(Cells(numVars + startRow + 3, 2), Cells(numVars + startRow + 3, 8)).Select
2394      With Selection.Borders(xlEdgeBottom)
2395          .LineStyle = xlContinuous
2396          .ColorIndex = 0
2397          .TintAndShade = 0
2398          .Weight = xlMedium
2399      End With
2400      With Selection.Font
2401          .Bold = True
2402          .ThemeColor = xlThemeColorLight2
2403      End With
          
          'Bold the variable table headings and make them blue as well as put a border around them
2404      Sheets(nameSheet).Range(Cells(startRow, 2), Cells(startRow, 8)).Select
2405      With Selection.Borders(xlEdgeBottom)
2406          .LineStyle = xlContinuous
2407          .ColorIndex = 0
2408          .TintAndShade = 0
2409          .Weight = xlMedium
2410      End With
2411      With Selection.Font
2412          .Bold = True
2413          .ThemeColor = xlThemeColorLight2
2414      End With
          
          'Bold the headings
2415      With Range("A:A").Font
2416          .Bold = True
2417      End With
          
2418      Cells(100, 100).Select
2419      Sheets(sheet).Select

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("COpenSolver", "FormatSensitivityTable") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Function findName(sheet As String, cell As String) As String
'Finds the name of a constraint or variable by finding the nearest strings to the left and above the cell and putting these together
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim NotFoundStringLeft As Boolean, NotFoundStringTop As Boolean
          Dim row As Long, col As Long, i As Long, j As Long
          Dim CellValue As String, LHSName As String, AboveName As String

2421      row = Sheets(sheet).Range(cell).row
2422      col = Sheets(sheet).Range(cell).Column
2423      i = col - 1
2424      j = row - 1
          
2425      NotFoundStringLeft = True
2426      NotFoundStringTop = True
          'Loop through to the left and above the cell to find the first non-numeric cell
2427      While (NotFoundStringLeft And i > 0) Or (NotFoundStringTop And j > 0)
              'Find the nearest name to the left of the variable or constraint if one exists
2428          If i > 0 And NotFoundStringLeft Then
2429              CellValue = Sheets(sheet).Cells(row, i)
                  'x = IsAmericanNumber(CellValue)
2430              If Not (IsNumeric(CellValue) Or (left(CellValue, 1) = "=") Or (CellValue = "")) Then
2431                  NotFoundStringLeft = False
2432              End If
2433              i = i - 1
2434          End If
              'Find the nearest name above the variable or constraint if it exists
2435          If j > 0 And NotFoundStringTop Then
2436              CellValue = Sheets(sheet).Cells(j, col)
2437              If Not (IsNumeric(CellValue) Or (left(CellValue, 1) = "=") Or (CellValue = "")) Then
2438                  NotFoundStringTop = False
2439              End If
2440              j = j - 1
2441          End If
2442      Wend
2443      LHSName = Cells(row, i + 1)
2444      AboveName = Cells(j + 1, col)
          
          'Put the names together
2445      If AboveName = "" Then
2446          findName = LHSName
2447      ElseIf LHSName = "" Then
2448          findName = AboveName
2449      Else
2450          findName = LHSName & " " & AboveName
2451      End If

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("COpenSolver", "findName") Then Resume
          RaiseError = True
          GoTo ExitFunction

End Function

' NOMAD functions
' Do not put error handling in these functions, there is already error-handling in the NOMAD DLL

Function updateVarOS(X As Variant)

          Dim i As Long, numVars As Long

          'set new variable values on sheet
2452      numVars = UBound(X)
2453      i = 1
          Dim AdjCell As Range
          ' If only one variable is returned, X is treated as a 1D array rather than 2D, so we need to access it
          ' differently.
2454      If numVars = 1 Then
2455          For Each AdjCell In AdjustableCells
2456              AdjCell.Value2 = X(i)
2457              i = i + 1
2458          Next AdjCell
2459      Else
2460          For Each AdjCell In AdjustableCells
2461              AdjCell.Value2 = X(i, 1)
2462              i = i + 1
2463          Next AdjCell
2464      End If
          
End Function

Function getValuesOS() As Variant
          Dim X As Variant, i As Long, j As Long, k As Long, conCountB As Long, numCons As Variant
2465      numCons = getNumConstraintsOS()
2466      ReDim X(1 To numCons(0), 1 To 1)
          '====NOMAD only does minimise so need to change objective if it is max====
          ' If no objective, just set a constant.
          ' TODO: fix this to set it based on amount of violation to hunt for feasibility
2467      If ObjRange Is Nothing Then
2468          X(1, 1) = 0
          ' If objective cell is error, report this directly to NOMAD. Attempting to manipulate it can cause errors
2469      ElseIf VarType(ObjRange.Value2) = vbError Then
2470          X(1, 1) = ObjRange.Value2
          'If objective sense is maximise then multiply by minus 1
2471      ElseIf ObjectiveSense = MaximiseObjective Then
2472          If ObjRange.Value2 <> 0 Then
2473              X(1, 1) = -1 * ObjRange.Value2 'objective value
2474          Else
2475              X(1, 1) = ObjRange.Value2
2476          End If
          'Else if objective sense is minimise leave it
2477      ElseIf ObjectiveSense = MinimiseObjective Then
2478          X(1, 1) = ObjRange.Value2
2479      ElseIf ObjectiveSense = TargetObjective Then
2480          X(1, 1) = Abs(ObjRange.Value2 - ObjectiveTargetValue)
2481      End If
2482      conCountB = 1 'keep a count of what constraint its up to including bounds
2483      k = 1 'keep a count of what constraint its up to not including bounds
          Dim row As Long, constraint As Long
2484      row = 1
          Dim CurrentLHSValues As Variant
          Dim CurrentRHSValues As Variant
2485      For constraint = 1 To NumConstraints
              ' Check to see what is different and add rows to sparsea
2486          If Not LHSRange(constraint) Is Nothing Then ' skip Binary and Integer constraints
                  ' Get current value(s) for LHS and RHS of this constraint off the sheet. LHS is always an array (even if 1x1)
2487              GetCurrentConstraintValues constraint, CurrentLHSValues, CurrentRHSValues
2488              If LHSType(constraint) = SolverInputType.MultiCellRange Then
2489                  For i = 1 To UBound(LHSOriginalValues(constraint), 1)
2490                      For j = 1 To UBound(LHSOriginalValues(constraint), 2)
2491                          If VariableBounds(conCountB) = False Then
2492                              If RHSType(constraint) <> SolverInputType.MultiCellRange Then
2493                                  SetConstraintValue X, k, CurrentRHSValues, CurrentLHSValues(i, j), Relation(conCountB)
2494                              ElseIf UBound(LHSOriginalValues(constraint), 1) = UBound(RHSOriginalValues(constraint), 1) Then
2495                                  SetConstraintValue X, k, CurrentRHSValues(i, j), CurrentLHSValues(i, j), Relation(conCountB)
2496                              Else
2497                                  SetConstraintValueMismatchedDims X, k, CurrentRHSValues, CurrentLHSValues, Relation(conCountB), i, j
2498                              End If
2499                          End If
2500                          conCountB = conCountB + 1
2501                      Next j
2502                  Next i
2503              Else
2504                  If VariableBounds(conCountB) = False Then
2505                      SetConstraintValue X, k, CurrentRHSValues, CurrentLHSValues(1, 1), Relation(conCountB)
2506                  End If
2507                  conCountB = conCountB + 1
2508              End If
2509          End If
2510      Next constraint
          
          'Get back new objective and difference between LHS and RHS values
2511      getValuesOS = X
End Function

Sub SetConstraintValue(ByRef ConstraintValues As Variant, ByRef k As Long, RHSValue As Variant, LHSValue As Variant, RelationType As Long)
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

Sub SetConstraintValueMismatchedDims(ByRef ConstraintValues As Variant, ByRef k As Long, RHSValues As Variant, LHSValues As Variant, RelationType As Long, i As Long, j As Long)
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

Function DifferenceOrError(Value1 As Variant, Value2 As Variant) As Variant
2532            If VarType(Value1) = vbError Then
2533                DifferenceOrError = Value1
2534            ElseIf VarType(Value2) = vbError Then
2535                DifferenceOrError = Value2
2536            Else
2537                DifferenceOrError = Value1 - Value2
2538            End If
End Function


Function getNumVariablesOS() As Variant
2539      getNumVariablesOS = AdjustableCells.Count
End Function

Function getNumConstraintsOS() As Variant
          'The number of constraints is actually the number of Objectives + Number of Constraints
          'Note: Bounds do not count as constraints and equalities count as 2 constraints
          Dim i As Long
          Dim X(0 To 1) As Double
2540      X(0) = 1
2541      For i = 1 To NumRows
2542          If VariableBounds(i) = False Then X(0) = X(0) + 1
2543          If Relation(i) = RelationEQ Then X(0) = X(0) + 1
2544      Next i

          'Number of objectives - NOMAD can do bi-objective
          'will take the first two values passed to it from getValuesOS as the objectives if doing biobj
          'Note: Currently OpenSolver can only do single objectives- will need to set up multi objectives yourself
2545      X(1) = 1 'number of objectives
2546      getNumConstraintsOS = X
End Function

Function getVariableDataOS() As Variant
          Dim numVars As Double
2547      numVars = AdjustableCells.Count
          Dim X() As Double
2548      ReDim X(0 To 4 * numVars - 1)
          Dim i As Long, j As Long
2549      For i = 0 To numVars - 1
2550          If AssumeNonNegativeVars Then
2551              X(2 * i) = 0
2552          Else
2553              X(2 * i) = -10000000000000#
2554          End If
2555          X(2 * i + 1) = 10000000000000#
2556      Next i
          
          ' Reindex bounded variables by relative address so that VarNames(var) can search the collection
          Dim LowerBoundedVariables As New Collection, c As Range
2557      For Each c In AdjustableCells
2558          If TestKeyExists(VarLowerBounds, c.Address) Then
2559              LowerBoundedVariables.Add VarLowerBounds(c.Address), c.Address(RowAbsolute:=False, ColumnAbsolute:=False)
2560          End If
2561      Next c

          Dim UpperBoundedVariables As New Collection
          For Each c In AdjustableCells
              If TestKeyExists(VarUpperBounds, c.Address) Then
                  UpperBoundedVariables.Add VarUpperBounds(c.Address), c.Address(RowAbsolute:=False, ColumnAbsolute:=False)
              End If
          Next c
          
          
          ' Apply bounds
          Dim var As Long
2562      For var = 1 To numVars
2563          If TestKeyExists(LowerBoundedVariables, VarNames(var)) Then
2564              X(2 * var - 2) = LowerBoundedVariables(VarNames(var))
2565          End If
              If TestKeyExists(UpperBoundedVariables, VarNames(var)) Then
                  X(2 * var - 1) = UpperBoundedVariables(VarNames(var))
              End If
2566      Next var

          'Get the starting point
          'Takes the points on the sheet and forces them between the bounds
2567      For j = 0 To numVars - 1
2568          If StartingPosition(j + 1) < X(2 * j) Then
2569              X(j + 2 * numVars) = X(2 * j)
2570          ElseIf StartingPosition(j + 1) > X(2 * j + 1) Then
2571              X(j + 2 * numVars) = X(2 * j + 1)
2572          Else
2573              X(j + 2 * numVars) = StartingPosition(j + 1)
2574          End If
2575      Next j
          
          'Get the variable type(real, int or bin)
2576      For i = 1 To numVars
          'initialise all variables as continuous
2577          X(i - 1 + 3 * numVars) = 1
2578      Next i
          Dim counter As Long, types As Variant
2579      counter = 2
2580      For Each types In Array(IntegerCellsRange, BinaryCellsRange)
2581          If Not types Is Nothing Then
2582              For Each c In types
2583                  For i = 1 To numVars
2584                      If VarNames(i) = c.Address(RowAbsolute:=False, ColumnAbsolute:=False) Then
2585                          X(i - 1 + 3 * numVars) = counter
2586                          If Not SolveRelaxedModel Then
                                  'Make bounds on integer and binary constraints integer
2587                              If X(2 * i - 2) > 0 Then
2588                                  X(2 * i - 2) = Application.WorksheetFunction.RoundUp(X(2 * i - 2), 0)
2589                              Else
2590                                  X(2 * i - 2) = Application.WorksheetFunction.RoundDown(X(2 * i - 2), 0)
2591                              End If
2592                              If X(2 * i - 1) > 0 Then
2593                                  X(2 * i - 1) = Application.WorksheetFunction.RoundDown(X(2 * i - 1), 0)
2594                              Else
2595                                  X(2 * i - 1) = Application.WorksheetFunction.RoundUp(X(2 * i - 1), 0)
2596                              End If
                                  'Make starting positions on integer and binary constraints integer
2597                              If X(i - 1 + 2 * numVars) < X(2 * i - 2) Then
2598                                  X(i - 1 + 2 * numVars) = X(2 * i - 2)
2599                              ElseIf X(i - 1 + 2 * numVars) > X(2 * i - 1) Then
2600                                  X(i - 1 + 2 * numVars) = X(2 * i - 1)
2601                              Else
2602                                  X(i - 1 + 2 * numVars) = Round(X(i - 1 + 2 * numVars))
2603                              End If
2604                          End If
2605                      End If
2606                  Next i
2607              Next c
2608          End If
2609          counter = counter + 1
2610      Next types
          
2611      getVariableDataOS = X
End Function

Function getOptionDataOS() As Variant
          Dim X() As Double
2612      ReDim X(0 To 2)
          Dim SolveOptions As SolveOptionsType
2613      GetSolveOptions sheet, SolveOptions
2614      X(0) = SolveOptions.MaxIterations
2615      X(1) = SolveOptions.MaxTime
2616      X(2) = SolveOptions.Precision
2617      getOptionDataOS = X
End Function

