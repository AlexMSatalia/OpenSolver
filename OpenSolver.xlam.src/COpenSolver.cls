VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COpenSolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' OpenSolver
' Copyright Andrew Mason 2010
' http://www.OpenSolver.org
' This software is distributed under the terms of the GNU General Public License
'
'
' This file is part of OpenSolver.
'
' OpenSolver is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' OpenSolver is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with OpenSolver.  If not, see <http://www.gnu.org/licenses/>.
'

' OpenSolver v0.8
'
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
'
' 2010.05.28: Modified code to allow a variable to be specified as both binary and integer, meaning it is binary
'
' v0.982: Changed NumVars and NumRows from Integer to Long to handle larger problems
' v1.0: Changed all constraint handling to using ranges to improve speed - Kathleen Gilbert 01.2011
' v1.1: Add linearity checker

#Const DEBUG_ = True

Option Explicit
Option Base 1

Const EPSILON = 0.000001
Const ZERO = 0.00000001

Dim sheetName As String ' Sheetname in quotes as used for prefixing
Dim RawSheetName As String ' Just the worksheet name
Dim RawWorkBookName As String ' Just the workbook name
Dim book As Workbook, sheet As Worksheet

' Data for our model. We work thru the original Excel 'constraints' and form model 'rows' from these
Public AdjustableCells As Range    ' A range we form containing all the adjustable (ie decision) variables
Dim BinaryCellsRange As Range   ' A range we form containing all the adjustable (ie decision) variables that are binary
Dim IntegerCellsRange As Range  ' A range we form containing all the adjustable (ie decision) variables that are integer (which may include binary variables)
Dim AssumeNonNegativeVars As Boolean   ' Excel stores ...!solver_neg=1 if Assume Linear Model is turned on, 2 if off (NB: Not 0!)

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

Dim ObjRange As Range

' Details for the model that we build.
Dim numVars As Long, NumIntVars As Long, NumBinVars As Long, NumConstraints As Long
Dim NumRows As Long
Dim ObjectiveSense As Integer
Dim ObjectiveTargetValue As Double  ' Used when the user wants the objective to achieve some stated target value
Dim VarNames() As String
Dim AdjCellsSheetIndex As Integer    ' The sheet index containing variables with no index specified in the CBC result file
Dim CostCoeffs() As Double, ObjectiveFunctionConstant As Double
Dim RHS() As Double ' The RHS values for each row in the LP
Dim Relation() As Integer ' For each row, we remember its relation (=, <=, >=)

'Details for the chosen solver
Public Solver As String 'Name of solver we will be using
Dim SolverString As String 'Name of solver for printing

' Details we need for non-linear solver
Dim VariableBounds() As Boolean
Dim StartingPosition() As Double
Dim SolveRelaxedModel As Boolean

' Details for sensitivity analysis
Dim FinalValue() As String
Dim ShadowPrice() As String
Dim IncreaseCon() As String
Dim DecreaseCon() As String
Dim FinalVarValue() As String
Dim ReducedCosts() As String
Dim IncreaseVar() As String
Dim DecreaseVar() As String
Dim VarCell() As String

' This is the A matrix for our Solver model
Dim SparseA() As CIndexedCoeffs

' We remember how to turn a row index in the LP back into a constraint index in the Excel list of constraints
' This array stores the first row in the LP associated with each Excel constraint
Dim MappingRowsToConstraints() As Long

' Our result as an enum and a string
Public SolveStatus As OpenSolverResult ' See Enum in OpenSolverModule; Optimal,Infeasible, etc
Public SolveStatusString As String    ' Eg Infeasible, Opimal etc
Public SolveStatusComment As String ' Eg constraint 1 cannot be satisified
Public LinearSolveStatus As LinearSolveResult ' Unsolved, etc
Public LinearSolveStatusString As String ' The string read back from the CBC file
Public LinearSolutionWasLoaded As Boolean ' True if we loaded a solution into the spreadsheet

Dim ModelStatus As Integer
Const ModelStatus_Unitialized = 0
Const ModelStatus_Built = 1

' The following are calculated differently; TimeTakenToBuildModel fails if we cross midnight!
Dim TimeTakenToBuildModel As Single, SecondsTakenToBuildModel As Long

' Support for doing a quick-solve where we only modify the RHS
Dim QuickSolve As CQuickSolve

' This is the A matrix for our Solver model
' Dim A() As Double
' See top of file for this

'Sub RunOpenSolver()
'    Dim Success As Boolean
'    Success = BuildModelFromSolverData
'    If Success Then SolveModel
'End Sub

' See http://www.jkp-ads.com/OfficeMarketPlaceNM-EN.asp for a tool to view the Solver's hidden names
' See https://www.coin-or.org/OS/publications/OSICS2009.pdf for accessing a COIN-OR solver
' See http://ideas.repec.org/a/jss/jstsof/14i04.html Calling LPSolve from Excel (by Samuel E. Buttrey)
' See http://peltiertech.com/Excel/SolverVBA.html for info on Solver and VBA
' http://www.sce.carleton.ca/faculty/chinneck/MPSreader/MPSobject.html for an MPS file reader/writer

' TODO: In an LP file, The maximum length for any name is 255. The maximum length of any line of input is 510.

Public Function GetNumRows() As Long
7050      GetNumRows = NumRows
End Function

Sub GetConstraintFromRow(row As Long, ByRef constraint As Long, ByRef instance As Long)
          ' Get the Excel constraint that goes with this LP row
          ' The incoming Constraint value will be either 1 or the constraint found on the last call to this
          ' We return the entry where MappingRowsToConstraints(Constraint) <= Row < MappingRowsToConstraints(Constraint+1)
          ' Note that MappingRowsToConstraints may contain sequences of the same value where the Excel constraint is 'bin' or 'int' and so has no rows
#If DEBUG_ Then
7060      Debug.Assert (row >= 1 And row <= NumRows)
7070      Debug.Assert (constraint >= 1 And constraint <= UBound(MappingRowsToConstraints))
#End If
7080      If MappingRowsToConstraints(constraint) > row Then constraint = 1 ' Start searching from the beginning (TODO: Bisection search)
7090      While row >= MappingRowsToConstraints(constraint + 1)
7100          constraint = constraint + 1
7110      Wend
7120      instance = row - MappingRowsToConstraints(constraint) + 1
#If DEBUG_ Then
7130      Debug.Assert (constraint >= 1 And constraint <= UBound(MappingRowsToConstraints) And row >= MappingRowsToConstraints(constraint) And row < MappingRowsToConstraints(constraint + 1))
#End If
End Sub

Sub GetArrayPosnFromConstraintInstance(constraint As Long, instance As Long, i As Long, j As Long)
          ' Get the row (x) and constraint that goes with this LP row
          ' The incoming Constraint value will be either 1 or the constraint found on the last call to this
          ' We return the entry where MappingRowsToConstraints(Constraint) <= Row < MappingRowsToConstraints(Constraint+1)
          ' Note that MappingRowsToConstraints may contain sequences of the same value where the Excel constraint is 'bin' or 'int' and so has no rows
7140      i = 1 + ((instance - 1) \ UBound(LHSOriginalValues(constraint), 2))
7150      j = 1 + ((instance - 1) Mod UBound(LHSOriginalValues(constraint), 2))
End Sub

Function CheckConstraintValuesAreNumeric(constraint As Long, LHSValues As Variant, RHSValues As Variant, ErrorMsg As String) As Boolean
          ' Check that the LHS and RHS values are doubles (not vbError, or vbString, or ??)
          ' vbEmpty (blank) cell values are permitted as VBA quietly turns them into 0's in expressions
7160      CheckConstraintValuesAreNumeric = False
          Dim i As Long, j As Long, instance As Long
          Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
          
          ' Check LHS first; this is always an array
7170      instance = 0
7180      For i = 1 To UBound(LHSValues, 1)
7190          For j = 1 To UBound(LHSValues, 2)
7200              instance = instance + 1
7210              If VarType(LHSValues(i, j)) = vbEmpty Then
                      ' This is an empty cell; VBA quietly turns this into 0 when used in an expression. We need no special handling of this, and don't report an error.
7220              ElseIf VarType(LHSValues(i, j)) = vbError Then
7230                  GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
7240                  ErrorMsg = "The cell '" & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "' in constraint '" & ConstraintSummary(constraint) & "' has a divide by zero error or has a problem with the function used ( eg #DIV/0! or #VALUE!). Please fix this and try again."
7250                  Exit Function
7260              ElseIf VarType(LHSValues(i, j)) <> vbDouble Then
                      ' It might be vbError, or vbString, or vbEmpty
7270                  GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
7280                  ErrorMsg = "The cell '" & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "' in constraint '" & ConstraintSummary(constraint) & "' has a value which is not numeric."
7290                  On Error Resume Next
                      ' Try to show the incorrect value; just cary on with the old ErrorMsg if this gives an error
7300                  ErrorMsg = "The cell '" & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "' in constraint '" & ConstraintSummary(constraint) & "' has the value '" & LHSValues(i, j) & "' which is not numeric."
7310                  Exit Function
7320              End If
7330          Next j
7340      Next i
          
7350      Select Case RHSType(constraint)
              Case SolverInputType.SingleCellRange, SolverInputType.Formula, SolverInputType.constant
7360              If VarType(RHSValues) = vbEmpty Then
                      ' This is an empty cell; VBA quietly turns this into 0 when used in an expression. We need no special handling of this, and don't report an error.
7370              ElseIf VarType(RHSValues) = vbError Then
7380                  GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
7390                  ErrorMsg = "The right hand side in constraint '" & ConstraintSummary(constraint) & "' has a divide by zero error or has a problem with the function used ( eg #DIV/0! or #VALUE!). Please fix this and try again."
7400                  Exit Function
7410              ElseIf VarType(RHSValues) <> vbDouble Then
                      ' It might be vbError, or vbString, or anything else?
7420                  GetConstraintInstanceData constraint, 1, LHSCellRange, RHSCellRange, RHSstring
7430                  ErrorMsg = "The right hand side in constraint '" & ConstraintSummary(constraint) & "' has a value which is not numeric."
                      ' Try to show the incorrect value; just cary on with the old ErrorMsg if this gives an error
7440                  On Error Resume Next
7450                  ErrorMsg = "The right hand side in constraint '" & ConstraintSummary(constraint) & "' has the value '" & RHSValues & "' which is not numeric."
7460                  Exit Function
7470              End If
7480          Case SolverInputType.MultiCellRange
7490              instance = 0
7500              For i = 1 To UBound(RHSValues, 1)
7510                  For j = 1 To UBound(RHSValues, 2)
7520                      instance = instance + 1
7530                      If VarType(RHSValues(i, j)) = vbEmpty Then
                              ' This is an empty cell; VBA quietly turns this into 0 when used in an expression. We need no special handling of this, and don't report an error.
7540                      ElseIf VarType(RHSValues) = vbError Then
7550                          GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
7560                          ErrorMsg = "The right hand side in constraint '" & ConstraintSummary(constraint) & "' has a divide by zero error or has a problem with the function used ( eg #DIV/0! or #VALUE!). Please fix this and try again."
7570                          Exit Function
7580                      ElseIf VarType(RHSValues(i, j)) <> vbDouble Then
                              ' It might be vbError, or vbString, or vbEmpty
7590                          GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
7600                          ErrorMsg = "The cell '" & RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "' in constraint '" & ConstraintSummary(constraint) & "' has a value which is not numeric."
7610                          On Error Resume Next
7620                          ErrorMsg = "The cell '" & RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "' in constraint '" & ConstraintSummary(constraint) & "' has the value '" & RHSValues(i, j) & "' which is not numeric."
7630                          Exit Function
7640                      End If
7650                  Next j
7660              Next i
7670      End Select
7680      CheckConstraintValuesAreNumeric = True
End Function

Sub GetCurrentConstraintValues(constraint As Long, LHSValues As Variant, RHSValues As Variant)
          ' Get the current values off the spreadsheet for the given constraint. LHS is always returned as a 2D array. RHS may be a single or array value
7690      If LHSType(constraint) = SolverInputType.SingleCellRange Then
              ' Create a 1x1 array
7700          LHSValues = Create1x1Array(LHSRange(constraint).Value2) ' LHSRange(Constraint).value
7710      Else ' LHSConstraintType = SolverInputType.MultiCellRange
7720          LHSValues = LHSRange(constraint).Value2    ' Copy all the values into a variant, which will then be an array of doubles; see http://support.microsoft.com/kb/213798
7730      End If

7740      If RHSType(constraint) = SolverInputType.SingleCellRange Then
7750          RHSValues = RHSRange(constraint).Value2  ' A variant containing just a single double
7760      ElseIf RHSType(constraint) = SolverInputType.MultiCellRange Then
7770          RHSValues = RHSRange(constraint).Value2   ' Copy all the values into a variant, which will then be an array of doubles (or an array of variants?)
7780      ElseIf RHSType(constraint) = SolverInputType.Formula Then
7790          RHSValues = ActiveSheet.Evaluate("=" & RHSFormula(constraint)) ' Will not throw an error, but may put an error value in RHSValues
7800      Else        'if the rhs value is a constant
7810          RHSValues = RHSOriginalValues(constraint)
7820      End If
End Sub

Sub GetConstraintInstanceData(constraint As Long, instance As Long, LHSCellRange As Range, RHSCellRange As Range, RHSstring As String)
      ' Get the information required to display a single instance of a constraint to the user
7830      Select Case LHSType(constraint)
              Case SolverInputType.SingleCellRange
7840              Set LHSCellRange = LHSRange(constraint)
7850          Case SolverInputType.MultiCellRange
7860              Set LHSCellRange = GetOneCellInRange(LHSRange(constraint), instance)
7870          Case SolverInputType.Formula
7880              Debug.Assert False ' should never happen
7890          Case SolverInputType.constant
7900              Debug.Assert False ' should never happen
7910      End Select
7920      Select Case RHSType(constraint)
              Case SolverInputType.SingleCellRange
7930              Set RHSCellRange = RHSRange(constraint)
7940              RHSstring = ""
7950          Case SolverInputType.MultiCellRange
7960              Set RHSCellRange = GetOneCellInRange(RHSRange(constraint), instance)
7970              RHSstring = ""
7980          Case SolverInputType.Formula
7990              Set RHSCellRange = Nothing
8000              RHSstring = RHSFormula(constraint)
8010          Case SolverInputType.constant
8020              Set RHSCellRange = Nothing
8030              RHSstring = CStr(RHSOriginalValues(constraint))
8040      End Select
End Sub

Function GetDecisionVariableValuesOffSheet() As Double()
          ' Return an array containing the current values of the decision variables on the sheet.
          ' This needs improving for speed
          ' Note that the decision variables are defined by a range that may have multiple areas
          '      but must be all on one sheet (even if not the ActiveSheet)
          Dim Vars() As Double, var As Long
8050      ReDim Vars(numVars)
8060      For var = 1 To numVars
8070          Vars(var) = AdjustableCells.Worksheet.Range(VarNames(var)).Value2
8080      Next var
8090      GetDecisionVariableValuesOffSheet = Vars
End Function
    
Function CalcObjFnValue(DecisionVariableValues() As Double) As Double
          Dim sum As Double, var As Long
8100      sum = ObjectiveFunctionConstant ' The value found for the obj function when all decision variables are 0
8110      For var = 1 To numVars
8120          sum = sum + (CostCoeffs(var) * DecisionVariableValues(var))
8130      Next var
8140      CalcObjFnValue = sum
End Function

Sub BuildModelFromSolverData(Optional LinearityCheckOffset As Double = 0)
          ' Build the model; throws an error if it fails
          
8150      SolveStatus = OpenSolverResult.Unsolved
8160      SolveStatusString = "Unsolved"
8170      SolveStatusComment = ""
8180      LinearSolveStatus = LinearSolveResult.Unsolved
8190      LinearSolveStatusString = "Unsolved"
8200      LinearSolutionWasLoaded = False
          
8210      ModelStatus = ModelStatus_Unitialized
          
          Dim errorPrefix As String
8220      errorPrefix = "OpenSolver Model Build"
8230      On Error GoTo errorHandler

          Dim StartTimeOfBuild As Single, StartDateTimeOfBuild As Date
          Dim EndTimeOfBuild As Single, EndDateTimeOfBuild As Date
8240      StartTimeOfBuild = Timer    ' Warning: This goes back to zero at midnight!
8250      StartDateTimeOfBuild = Now()
          
8260      CheckWorksheetAvailable SuppressDialogs:=True, ThrowError:=True ' Will throw if there is no workbook or worksheet

          ' We trap the Escape key which does an onerror
          'xlDisabled = 0 'totally disables Esc / Ctrl-Break / Command-Period
          'xlInterrupt = 1 'go to debug
          'xlErrorHandler = 2 'go to error handler
          'Trappable error is #18
8270      Application.EnableCancelKey = xlErrorHandler
8280      On Error GoTo errorHandler
          
8290      On Error Resume Next
8300      Set sheet = ActiveWorkbook.ActiveSheet
          ' This error checking is not really needed as it also occurs in CheckWorkSheetAvailable
8310      If Err.Number <> 0 Then
8320          On Error GoTo errorHandler
8330          Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:="Unable to access the active worksheet."
8340      End If
8350      On Error GoTo errorHandler
8360      RawSheetName = sheet.Name
8370      sheetName = "'" & Replace(RawSheetName, "'", "''") & "'!"   ' Must include quotes to handle sheet names with spaces, & double any ' when we quote the name
          
8380      Set book = ActiveWorkbook
8390      RawWorkBookName = book.Name
        
          '----------------------------------------------------------------------------------------------------------
          'Find the users preferred solver
8400      If GetNameValueIfExists(ActiveWorkbook, sheetName & "OpenSolver_ChosenSolver", Solver) Then
              If Solver Like "NeosCBC" Then
                  Solver = "NeosCBC"
                  SolverString = "NEOSCBC"
8410          ElseIf Solver Like "*bc*" Or Solver Like "*BC*" Then
8420              Solver = "cbc.exe"
8430              SolverString = "CBC"
8440          ElseIf Solver Like "*urobi" Then
8450              Solver = "gurobi.bat"
8460              SolverString = "Gurobi"
8470          ElseIf Solver Like "NOMAD*" Then
8480              Solver = "Nomad"
8490              SolverString = "NOMAD"
              '---------Need to add in cplex solver---------------------------------------
8500          ElseIf Solver Like "*plex" Then
8510              Solver = "cplex.exe"
8520              SolverString = "CPLEX"
          '          ElseIf Solver = "NeosCBC" Then
          '              solver = "qsopt_ex"
              ElseIf Solver Like "*Cou" Then
                  Solver = "NEOSCou"
                  SolverString = "Couenne"
                  GoTo ExitSub
              ElseIf Solver Like "*Bon" Then
                  Solver = "NEOSBon"
                  SolverString = "Bonmin"
                  GoTo ExitSub
              ElseIf Solver Like "PuLP" Then
                  Solver = "PuLP"
                  SolverString = "PuLP"
                  GoTo ExitSub
8530          Else
8540            Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:="OpenSolver did not recognise the solver you have picked. Please pick another one and try again."
8550          End If
          Else
              Solver = "cbc.exe"
              SolverString = "CBC"
8560      End If

          ' We check to see if a model exists by getting the adjustable cells
          ' TODO: For some unknown reason, this fails if the sheet name as a ! in it, even when it is quoted.
          Dim n As Name
8570      On Error Resume Next
8580      Set n = Names(sheetName & "solver_adj")
          
8590      If Err.Number <> 0 Then
8600          On Error GoTo errorHandler
8610          Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:="No Solver model was found on sheet " & RawSheetName
8620      End If
          
          ' Change to manual calculation mode
          Dim oldCalculationMode As Integer
8630      oldCalculationMode = Application.Calculation
8640      Application.Calculation = xlCalculationManual
          
          Dim i As Long, j As Long

8650      Application.ScreenUpdating = False
8660      Application.Cursor = xlWait

          '--------------------------------------------------------------------------
          ' ADJUSTABLE CELLS SETUP and Setting of Initial Decision Var Cell values
          '--------------------------------------------------------------------------
8670      On Error Resume Next
          ' Get the range stored in the (hidden) Solver name solver_adj
          Dim rTemp As Range
8680      Set rTemp = Range(sheetName & "solver_adj")
          ' Catch any error related to either the model being incomplete, or invalid
          ' entry for adjustable cells.
8690      If Err.Number <> 0 Then
8700          On Error GoTo errorHandler
8710          Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:="A model was found on the sheet " & RawSheetName & " but the decision variable cells (" & n & ") could not be interpreted. Please redefine the decision variable cells, and try again."
8720      End If
          ' Remove any overlap
8730      On Error GoTo errorHandler
8740      Set AdjustableCells = RemoveRangeOverlap(rTemp)
          ' Handle any merged ranges
          Dim BadCell As Range
8750      If Not CheckRangeContainsNoAmbiguousMergedCells(AdjustableCells, BadCell) Then
8760          On Error GoTo errorHandler
8770          Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:="This model contains a decision variable " & BadCell.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & " that is inaccessible as it is within the merged range " & BadCell.MergeArea.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & ". Please redefine the decision variable cells, and try again."
8780          GoTo ExitSub
8790      End If
          ' Check we have any cells after all this
8800      numVars = AdjustableCells.Count
8810      If numVars < 1 Then
8820          On Error GoTo errorHandler
8830          Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:="There are no decision variable cells in this model, and so the model cannot be solved."
8840      End If
          ' Find out which worksheet has the majority of our decision variables
          'Dim CellCount() As Long, area As Range
          'AdjCellsSheetIndex = 1
          'ReDim CellCount(ActiveWorkbook.Worksheets.Count)
          'For Each area In AdjustableCells.Areas
          '    CellCount(area.Worksheet.Index) = CellCount(area.Worksheet.Index) + area.Count
          '    If CellCount(area.Worksheet.Index) > CellCount(AdjCellsSheetIndex) Then AdjCellsSheetIndex = area.Worksheet.Index
          'Next area
          ' Get names for all the variables, creating names that distinguish cells on different sheets
8850      ReDim VarNames(numVars)
8860      ReDim CostCoeffs(numVars) As Double
8870      If Solver = "Nomad" Then
8880          ReDim StartingPosition(numVars)
8890      End If
          Dim AdjCell As Range
8900      i = 0
8910      For Each AdjCell In AdjustableCells
8920          i = i + 1
8930          VarNames(i) = AdjCell.Address(RowAbsolute:=False, ColumnAbsolute:=False) ' eg A1
8940          If Solver = "Nomad" Then
8950              StartingPosition(i) = AdjCell.value
8960          End If
8970      Next AdjCell
          ' Set all the adjustable cells to zero (or a user-specified value) (For linear problems)
8980      AdjustableCells.Value2 = LinearityCheckOffset
          
          '--------------------------------------------------------------------------
          ' INITIAL SHEET CALCULATION (using new decision var values set just above)
          '--------------------------------------------------------------------------
          ' We expect, after this, that our calculation state will be xlDone. However, for large s/sheets, the
          ' state may be xlCalculating or xlPending even after a calculation.
          ' if this happens, then we cannot be sure that Application.Calculate will work, so we warn the user.
          ' TODO: Currently we require the calculation to complete before we continue. Should we make this more flexible?
9000      If Not ForceCalculate("Warning: The initial worksheet calculation did not complete, and so the model may not be generated correctly. Would you like to try again?") Then GoTo ExitSub
          'End If
          '--------------------------------------------------------------------------
          ' OBJECTIVE SETUP
          '--------------------------------------------------------------------------
          ' Get the objective sense
          Dim temp As Integer
9010      ObjectiveSense = UnknownObjectiveSense
9020      If GetNamedIntegerIfExists(book, sheetName & "solver_typ", temp) Then ObjectiveSense = temp
9030      ObjectiveTargetValue = 0
9040      If ObjectiveSense = TargetObjective Then GetNamedNumericValueIfExists book, sheetName & "solver_val", ObjectiveTargetValue    ' Target will be 0 if this fails
          ' Get and check the objective function
          Dim isRangeObj As Boolean, valObj As Double, ObjRefersToError As Boolean
          Dim ObjRefersToFormula As Boolean, sRefersToObj As String, objIsMissing As Boolean
9050      GetNameAsValueOrRange book, sheetName & "solver_opt", objIsMissing, isRangeObj, ObjRange, ObjRefersToFormula, ObjRefersToError, sRefersToObj, valObj
          Dim objValue As Double
          ' Must have an objective
9060      If objIsMissing Then
9070          GoTo ModelOptions
9080          On Error GoTo errorHandler
9090          Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:="OpenSolver cannot find any objective ('solver_opt' is missing). Please fix the objective, and try again."
9100          GoTo ExitSub
9110      End If
          ' Objective is corrupted somehow
9120      If ObjRefersToError Then
9130          On Error GoTo errorHandler
9140          Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:="The objective is marked #REF!, indicating this cell has been deleted. Please fix the objective, and try again."
9150          GoTo ExitSub
9160      End If
          ' Objective has a value that is not a number
          Dim X As Variant
9170      If VarType(ObjRange.Value2) <> vbDouble Then
9180          On Error GoTo errorHandler
9190          If VarType(ObjRange.Value2) = vbError Then
9200              Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:="The objective cell appears to contain an error. This could have occurred if there is a divide by zero error or if you have used the wrong function (eg #DIV/0! or #VALUE!). Please fix this, and try again."
9210          Else
9220              Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:="The objective cell does not appear to contain a numeric value. Please fix this, and try again."
9230          End If
9240          GoTo ExitSub
9250      End If
          ' Store the value currently in the

9260      ObjRange.Calculate ' This does not always seem to work in that it doesn't track back predecessors; we now do a full re-calc above anyway
9270      objValue = ObjRange.Value2
9280      ObjectiveFunctionConstant = ObjRange.Value2

ModelOptions:

          ' 4/3/2011: We now correctly handle "Assume Non Negative". Excel adds a zero lower bound to any variable which does not have any other lower bound specified
          Dim s As String
9290      AssumeNonNegativeVars = False
9300      If GetNameValueIfExists(ActiveWorkbook, sheetName & "solver_neg", s) Then AssumeNonNegativeVars = s = "1"

          Dim succeeded As Boolean
9310      If Solver <> "Nomad" Then
9320          On Error GoTo errorHandler
9330          Call BuildModelOptions(s, succeeded)
9340          If succeeded = False Then GoTo ExitSub
9350      End If
          
          '--------------------------------------------------------------------------
          ' CONSTRAINTS SETUP
          '--------------------------------------------------------------------------
          ' Count the correct number of constraints - not necessarily the number of
          ' rows though.
          ' NB: Solver leaves around old constraints - the name <sheet>!solver_num
          ' gives the correct number of constraints (eg "=4")
9360      NumConstraints = Val(Mid(Names(sheetName & "solver_num"), 2))
          ' About to process the constraints, which takes a while, so provide some
          ' user feedback!
9370      Application.StatusBar = "OpenSolver: Setting Up Problem... " & numVars & " vars, " & NumConstraints & " Solver constraints"
          'Set up arrays to be the correct size. We allow no constraints!
9380      If NumConstraints > 0 Then
9390          ReDim LHSOriginalValues(NumConstraints) As Variant
9400          ReDim RHSOriginalValues(NumConstraints) As Variant
9410          ReDim RHSType(NumConstraints) As SolverInputType
9420          ReDim LHSType(NumConstraints) As SolverInputType
9430          ReDim LHSRange(NumConstraints) As Range
9440          ReDim RHSRange(NumConstraints) As Range
9450          ReDim RHSFormula(NumConstraints) As String
9460          ReDim ConstraintSummary(NumConstraints) As String
9470          ReDim MappingRowsToConstraints(NumConstraints + 1) As Long
9480      End If
          ' Reset number of rows in the LP model
9490      NumRows = 0
          Dim constraint As Long, row As Long
9500      row = 1
          ' Iterate through each solver constraint
9510      For constraint = 1 To NumConstraints
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
              Dim isRangeLHS As Boolean, valLHS As Double, rangeLHS As Range, LHSRefersToError As Boolean
              Dim LHSRefersToFormula As Boolean, sRefersToLHS As String, LHSisMissing As Boolean
              
9520          GetNameAsValueOrRange book, sheetName & "solver_lhs" & constraint, LHSisMissing, isRangeLHS, rangeLHS, LHSRefersToFormula, LHSRefersToError, sRefersToLHS, valLHS
              ' Must have a left hand side defined
9530          If LHSisMissing Then
9540              On Error GoTo errorHandler
9550              Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:="The left hand side for a constraint does not appear to be defined ('solver_lhs" & constraint & " is missing). Please fix this, and try again."
9560          End If
              ' Must be valid
9570          If LHSRefersToError Then
9580              On Error GoTo errorHandler
9590              Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:="The constraints reference cells marked #REF!, indicating these cells have been deleted. Please fix these constraints, and try again."
9600          End If
              ' LHSs must be ranges
9610          If Not isRangeLHS Then
9620              On Error GoTo errorHandler
9630              Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:="A constraint was entered with a left hand side (" & sRefersToLHS & ") that is not a range. Please update the constraint, and try again."
9640          End If
              
              '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
              ' RELATIONSHIP
              '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
              Dim rel As Integer
              ' Either <=, = or >=, or int or bin (or 'alldif' in Excel 2010)
9650          rel = Val(Mid(Names(sheetName & "solver_rel" & constraint), 2))
                      
              Dim RHSCount As Long, RowCount As Long
9660          RowCount = rangeLHS.Count
              
              ' If its integer or binary, there is no right-hand-side
9670          If rel = RelationINT Or rel = RelationBIN Then
                  ' Make the LHS variables integer or binary
                  ' Check that they haven't set as integer or binary something that
                  ' is not a decision variable...
                  Dim intersection As Range
9680              Set intersection = Intersect(rangeLHS, AdjustableCells)
9690              If intersection Is Nothing Then
                      ' The cells are not all decision variables
9700                  On Error GoTo errorHandler
9710                  Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:="A range of cells are specified as bin or int that are not decision variables. OpenSolver does not support this yet."
9720              End If
9730              If intersection.Count = rangeLHS.Count Then
9740                  If rel = RelationINT Then
9750                      Set IntegerCellsRange = ProperUnion(IntegerCellsRange, rangeLHS)
9760                  Else 'if the relation is binary
9770                      Set BinaryCellsRange = ProperUnion(BinaryCellsRange, rangeLHS)
9780                  End If
9790              Else
                      ' The cells are not all decision variables
9800                  On Error GoTo errorHandler
9810                  Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:="Unable to find a cell specified as bin or int in the decision variables."
9820              End If
                  ' This constraint does not relate to any new rows, so mark it
                  ' with a dummy entry. NumRows+1 = same value as we get for the
                  ' next entry, so this refers to no rows.
9830              MappingRowsToConstraints(constraint) = NumRows + 1
              
              ' If the constraint is not int or bin...
              '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
              ' RIGHT HAND SIDE
              '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
9840          Else
                  ' ... then the constraint is a full equation with a RHS
                  Dim isRangeRHS As Boolean, valRHS As Double, rangeRHS As Range, RHSRefersToError As Boolean
                  Dim RHSRefersToFormula As Boolean, sRefersToRHS As String, RHSisMissing As Boolean
                  
9850              GetNameAsValueOrRange book, sheetName & "solver_rhs" & constraint, RHSisMissing, isRangeRHS, rangeRHS, RHSRefersToFormula, RHSRefersToError, sRefersToRHS, valRHS
                  ' Must have a right hand side defined
9860              If RHSisMissing Then
9870                  On Error GoTo errorHandler
9880                  Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:="The right hand side for a constraint does not appear to be defined ('solver_rhs" & constraint & " is missing). Please fix this, and try again."
9890              End If
                  ' Must be valid
9900              If RHSRefersToError Then
9910                  On Error GoTo errorHandler
9920                  Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:="The constraints reference cells marked #REF!, indicating these cells have been deleted. Please fix these constraints, and try again."
9930              End If
                  
                  ' Track the constraints
                  Dim summary As String
9940              summary = ""
9950              summary = summary & Mid(Names(sheetName & "solver_lhs" & constraint), 2)
9960              summary = summary & " " & SolverRelationAsString(rel) & " "
9970              summary = summary & Mid(Names(sheetName & "solver_rhs" & constraint), 2)
9980              ConstraintSummary(constraint) = summary
                  
                  ' Don't allow 'alldiff' from Excel 2010 or any other wierd things
9990              If rel <> RelationLE And rel <> RelationEQ And rel <> RelationGE Then
10000                 On Error GoTo errorHandler
10010                 Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:="Unrecognised relationship for constraint " & ConstraintSummary(constraint)
10020             End If

                  ' Check we have a compatible constraint system.
                  ' Solver should have checked this already, but just to make sure...
10030             RHSCount = 1
10040             If isRangeRHS Then RHSCount = rangeRHS.Count
10050             If (RowCount <> 1 And RHSCount <> 1) And (RowCount <> RHSCount) Then
10060                 On Error GoTo errorHandler
10070                 Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:="The constraint '" & summary & "' has a different cell count on the left and the right. The model cannot be built."
10080             End If
                  
                  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                  ' STORE CONSTRAINT IN MEMORY
                  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                  ' Each constraint is stored as one row of LHSOriginalvalues and
                  ' RHSOriginalvalues which essentially become (variant) arrays of
                  ' variants. LHSType and RHSType tell us what is stored in each row
                  ' of LHSOriginalvalues and RHSOriginalvalues
                  
                  ' Left hand side:
10090             Set LHSRange(constraint) = rangeLHS
10100             If RowCount = 1 Then
10110                 LHSType(constraint) = SolverInputType.SingleCellRange
10120                 LHSOriginalValues(constraint) = Create1x1Array(rangeLHS.Value2)  ' Create a 1x1 array for the single cell range
10130             Else
10140                 LHSType(constraint) = SolverInputType.MultiCellRange
10150                 LHSOriginalValues(constraint) = rangeLHS.Value2 ' Copy the entire range over into an array of variants
10160             End If
                  
                  ' Right hand side:
                  ' If it is a range, it is much like LHS
10170             If isRangeRHS Then
10180                 Set RHSRange(constraint) = rangeRHS
10190                 RHSType(constraint) = IIf(rangeRHS.Count = 1, SolverInputType.SingleCellRange, SolverInputType.MultiCellRange)
10200                 RHSOriginalValues(constraint) = rangeRHS.Value2  ' Copy the entire range over into an array of variants
                  ' If its a formula, store the formula internally as a string
10210             ElseIf RHSRefersToFormula Then
10220                 RHSFormula(constraint) = sRefersToRHS
10230                 RHSType(constraint) = SolverInputType.Formula
10240                 RHSOriginalValues(constraint) = ActiveSheet.Evaluate("=" & sRefersToRHS)
                  ' Its a constant
10250             Else
10260                 RHSType(constraint) = SolverInputType.constant
10270                 RHSOriginalValues(constraint) = valRHS
10280             End If
                  
                  ' Store relations for these rows, keep track of how many rows we've added
10290             ReDim Preserve Relation(NumRows + rangeLHS.Count) As Integer
10300             For i = 1 To rangeLHS.Count
10310                 Relation(row + i - 1) = rel
10320             Next i
10330             row = row + rangeLHS.Count
                  ' Store the first LP row associated with this Excel constraint
10340             MappingRowsToConstraints(constraint) = NumRows + 1
10350             NumRows = NumRows + RowCount
                  
                  ' Check that all our values are in fact numbers
                  Dim ErrorMsg As String
10360             If Not CheckConstraintValuesAreNumeric(constraint, LHSOriginalValues(constraint), RHSOriginalValues(constraint), ErrorMsg) Then
10370                 On Error GoTo errorHandler
10380                 Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:=ErrorMsg
10390             End If
10400         End If
10410     Next constraint
          
          '--------------------------------------------------------------------------
          ' BUILD THE A MATRIX
          '--------------------------------------------------------------------------
          ' Initialise a sparse matrix
10420     If NumRows > 0 Then ReDim Preserve SparseA(NumRows) As CIndexedCoeffs
10430     For row = 1 To NumRows
10440         Set SparseA(row) = New CIndexedCoeffs
10450     Next row
10460     If NumRows > 0 Then ReDim RHS(NumRows) As Double
          
          '=======================================================================================================================================
          'TODO: Want to split up BuildConstraintMatrix so doesnt need to build when using Nomad. However Nomad is using the Sparse matrix to know
          'which constraints are bounds on the variables as these are treated differently
          '=======================================================================================================================================
          'Set up the constraint matrix for solving linear problems
10470     Call BuildConstraintMatrix(objIsMissing, objValue, SparseA, succeeded)
10480     If succeeded = False Then GoTo ExitSub


10490     ModelStatus = ModelStatus_Built
10500     EndTimeOfBuild = Timer
10510     EndDateTimeOfBuild = Now()
10520     SecondsTakenToBuildModel = DateDiff("s", StartDateTimeOfBuild, EndDateTimeOfBuild)
10530     TimeTakenToBuildModel = EndTimeOfBuild - StartTimeOfBuild

ExitSub:
          ' This code needs to be repeated below under ErrorExit:
10540     Application.Cursor = xlDefault
10550     Application.StatusBar = False ' Resume normal status bar behaviour
10560     Application.ScreenUpdating = True
10570     Application.Calculation = oldCalculationMode
10580     Exit Sub
          
errorHandler:
          ' We only trap Escape (Err.Number=18) here; all other errors are passed back to the caller.
          ' Save error message
          Dim ErrorNumber As Long, ErrorDescription As String, ErrorSource As String
10590     ErrorNumber = Err.Number
10600     ErrorDescription = Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
10610     ErrorSource = Err.Source
10620     If Err.Number = 18 Then
10630         If MsgBox("You have pressed the Escape key. Do you wish to cancel?", _
                         vbCritical + vbYesNo + vbDefaultButton1, _
                         "OpenSolver: User Interrupt Occured...") = vbNo Then
10640             Resume 'continue on from where error occured
10650         Else
                  ' Raise a "user cancelled" error. We cannot use Raise, as that exits immediately without going thru our code below
10660             ErrorNumber = OpenSolver_UserCancelledError
10670             ErrorSource = errorPrefix
10680             ErrorDescription = "Model building cancelled by user."
10690         End If
10700     End If
          
ErrorExit:
          ' Exit, raising an error; none of the following change the Err.Number etc, but we saved them above just in case...
10710     Application.Cursor = xlDefault
10720     Application.StatusBar = False ' Resume normal status bar behaviour
10730     Application.ScreenUpdating = True
10740     Application.Calculation = oldCalculationMode
10750     Err.Raise ErrorNumber, ErrorSource, ErrorDescription

      ' Debugging resume; never reached via code
10760     Resume  ' Used for manual trying again after an error
End Sub

Sub BuildModelOptions(s As String, succeeded As Boolean)
10770     succeeded = False
          Dim errorPrefix As String
10780     errorPrefix = "OpenSolver Model Build"
          
          '--------------------------------------------------------------------------
          ' MODEL OPTIONS
          '--------------------------------------------------------------------------
          
          ' Check that the user has Assume Linear Model checked.
          ' In Excel 2010, Assume Linear Model has been replaced by an Engine Type as Simplex.
10790     Dim AssumeLinearModel As Boolean ' Excel stores ...!solver_lin=1 if Assume Linear Model is turned on, 2 if off (NB: Not 0!)
          Dim SimplexEngineSelected As Boolean '2010 excel - stores the type of engine used

          ' If we are in Excel 2010, and the user has set "Solve Relaxtion", then we report an error
          Dim IngoreIntConstraints As Boolean
10800     IngoreIntConstraints = False
10810     If Application.Version = "14.0" And GetNameValueIfExists(ActiveWorkbook, sheetName & "solver_rlx", s) Then
10820         IngoreIntConstraints = s = "1"
10830     End If
10840     If IngoreIntConstraints = True Then
              'On Error GoTo errorHandler
10850         Err.Raise Number:=OpenSolver_BuildError, Source:=errorPrefix, Description:="You have the 'ignore integer constraints' option turned on. Please turn this off and start OpenSolver again or choose OpenSolver's solve relaxed model option"
10860     End If

          ' Get "Assume Linear Model option"; this is an Excel 2007 option; may not be present in Excel 2010 (unless opening an old file)
10870     AssumeLinearModel = False
10880     If GetNameValueIfExists(ActiveWorkbook, sheetName & "solver_lin", s) Then AssumeLinearModel = s = "1"

          ' Get the selected optimisation Engine, an Excel 2010 option (eg Nonlinear Solver, GA, LP). Will not be present in models built using Excel 2007
10890     SimplexEngineSelected = False ' a non-linear solver
10900     If GetNameValueIfExists(ActiveWorkbook, sheetName & "solver_eng", s) Then SimplexEngineSelected = s = "2"

10910     If (Not SimplexEngineSelected And Not AssumeLinearModel) Then
              ' User has not told us to use the Simplex method...
10920         If Application.Version <= 12# Then ' Excel 2007
10930              CustomForm.TextBox1.Caption = "OpenSolver assumes the model is linear. However, Assume Linear Model is not turned on. Would you like to continue anyway (and solve the problem as if Assume Linear Model was turned on)?"
10940              CustomForm.CheckBox1.Caption = "Turn on Assume Linear Model"
10950         Else   'if its excel 2010
10960              CustomForm.TextBox1.Caption = "OpenSolver always uses the simplex algorithm. However, the Solving Method is not set as Simplex. Would you like to continue anyway (and solve the problem using the Simplex method)?"
10970              CustomForm.CheckBox1.Caption = "Set the Solving Method to Simplex"
10980         End If
10990         CustomForm.Tag = ""
11000         CustomForm.Show
11010         If CustomForm.Tag = "Cancelled" Then
11020             Exit Sub
11030         ElseIf CustomForm.CheckBox1.value = True Then
11040             AssumeLinearModel = True
11050             If Application.Version <= 12# Then   ' Excel 2007, or earlier
11060                 SetSolverNameOnSheet "lin", "=1"
11070             Else 'excel 2010
11080                 SetSolverNameOnSheet "eng", "=2"
11090             End If
11100         Else    'if the user continues without correcting issues warn them of the dangers
11110             If MsgBox("Warning are you sure you want to continue without changing your model's settings?", vbCritical + vbYesNo + vbDefaultButton1, "OpenSolver Warning") = vbNo Then
11120                 CustomForm.Show
11130                 If CustomForm.Tag = "Cancelled" Then
11140                     Exit Sub
11150                 ElseIf CustomForm.CheckBox1.value = True Then
11160                     AssumeLinearModel = True
11170                     If Application.Version = 12# Then   ' Excel 2007
11180                         SetSolverNameOnSheet "lin", "=1"
11190                     Else 'excel 2010
11200                         SetSolverNameOnSheet "eng", "=2"
11210                     End If
11220                 End If
11230             End If
11240         End If
11250     End If
11260     succeeded = True
End Sub

Sub BuildConstraintMatrix(objIsMissing As Boolean, objValue As Double, SparseA() As CIndexedCoeffs, succeeded As Boolean)
          On Error GoTo constraintErrorHandler
          Dim errorPrefix As String
          errorPrefix = "OpenSolver Model Build"
11270     succeeded = False
          
          '--------------------------------------------------------------------------
          ' FILL A MATRIX BY CHANGING ADJUSTABLE CELLS
          '--------------------------------------------------------------------------
          ' Incrementing each decision variable by one to see what constraints are
          ' dependent on it and by what amount
          Dim var As Long
          Dim AdjCell As Range
          
          ' Add constants to Target
          If ObjectiveSense = TargetObjective Then
              ' Target value needs to be incremented by any constants in the objective
              ObjectiveTargetValue = ObjectiveTargetValue - objValue
          End If
          
11280     var = 0
11290     For Each AdjCell In AdjustableCells
              ' Count how many variables we've processed, let the user know
              ' that the program hasn't crashed!
11300         var = var + 1
11310         If var Mod 10 = 0 Then Application.StatusBar = "OpenSolver: Setting Up Problem... " & var & "/" & numVars & " vars, " & NumRows & " rows."

              ' Increment the value of this cell to find its coefficients
              ' Because its all linear, we can simply check the change in the objective/constraints
11320         AdjCell.Value2 = AdjCell.Value2 + 1

              ' Recalculate the spreadsheet
11330         If Not ForceCalculate("Warning: The worksheet calculation did not complete, and so the model may not be generated correctly. Would you like to retry?") Then Exit Sub

              ' Easy part first - the objective function value change
11340         If Not objIsMissing Then
11350             CostCoeffs(var) = ObjRange.Value2 - objValue
11360         End If
              ' Now the hard (slow) bit - the constraint changes
              Dim CurrentLHSValues As Variant
              Dim CurrentRHSValues As Variant
              Dim row As Long, constraint As Long
11370         row = 1
11380         For constraint = 1 To NumConstraints

                  ' Check to see what is different and add rows to sparsea
11390             If Not LHSRange(constraint) Is Nothing Then ' skip Binary and Integer constraints

                      Dim OriginalLHSValues As Variant
                      Dim OriginalRHSValues As Variant
11400                 OriginalLHSValues = LHSOriginalValues(constraint)    'accessing what has been stored for each constraint to find the coefficents
11410                 OriginalRHSValues = RHSOriginalValues(constraint)

                      ' Get current value(s) for LHS and RHS of this constraint off the sheet. LHS is always an array (even if 1x1)
11420                 GetCurrentConstraintValues constraint, CurrentLHSValues, CurrentRHSValues

                      'If LHSType(Constraint) = SolverInputType.SingleCellRange Then
                      '    If RHStype(Constraint) = SolverInputType.SingleCellRange Or RHStype(Constraint) = SolverInputType.Formula Then
                      '            coeff = (CurrentLHSValues - OriginalLHSValues) - (CurrentRHSValues - OriginalRHSValues)
                      '            If coeff <> 0 Then SparseA(Row).Add var, coeff
                      '            Row = Row + 1
                      '    Else    'If the RHS is constant it doesnt need to be considered
                      '            coeff = (CurrentLHSValues - OriginalLHSValues)
                      '            If coeff <> 0 Then SparseA(Row).Add var, coeff
                      '            Row = Row + 1
                      '    End If
                      'Else    'if the lhs is a multicell range
                      'TODO Put in better error trapping for errors that occur on the s/sheet

                      Dim coeff As Double, i As Long, j As Long
11430                 For i = 1 To UBound(OriginalLHSValues, 1) ' number of rows
11440                     For j = 1 To UBound(OriginalLHSValues, 2) ' number of columns
11450                         If RHSType(constraint) = SolverInputType.MultiCellRange Then
                                  '-----------------------------------------------------------
                                  'Making it work for column LHS with row RHS and vice versa
                                  '-----------------------------------------------------------
11460                             If UBound(OriginalLHSValues, 1) = UBound(OriginalRHSValues, 1) Then
11470                                 coeff = (CurrentLHSValues(i, j) - OriginalLHSValues(i, j)) - (CurrentRHSValues(i, j) - OriginalRHSValues(i, j))
11480                             Else
11490                                 coeff = (CurrentLHSValues(i, j) - OriginalLHSValues(i, j)) - (CurrentRHSValues(j, i) - OriginalRHSValues(j, i))
11500                             End If
11510                         ElseIf RHSType(constraint) = SolverInputType.Formula Or RHSType(constraint) = SolverInputType.SingleCellRange Then
11521                             coeff = (CurrentLHSValues(i, j) - OriginalLHSValues(i, j)) - (CurrentRHSValues - OriginalRHSValues)
11530                         Else    ' if the rhs is a constant value
11540                             coeff = (CurrentLHSValues(i, j) - OriginalLHSValues(i, j))
11550                         End If
11560                         If coeff <> 0 Then SparseA(row).Add var, coeff
11570                         row = row + 1
11580                     Next j
11590                 Next i
                      ' End If
11600             End If
11610         Next constraint
11620         AdjCell.Value2 = AdjCell.Value2 - 1
11630     Next AdjCell
          
11640     If (NumConstraints > 0) Then
11650         MappingRowsToConstraints(NumConstraints + 1) = NumRows + 1  ' An 'end of data' entry
11660     End If

          'Create the rhs vector
          ' TODO Handle errors in the original spreadsheet values
11670     row = 1
11680     For constraint = 1 To NumConstraints
11690         If Not LHSRange(constraint) Is Nothing Then
11700             If RHSType(constraint) = SolverInputType.MultiCellRange Then 'does the constraint refer to more than one cell
11710                 For i = 1 To UBound(LHSOriginalValues(constraint), 1)
11720                     For j = 1 To UBound(LHSOriginalValues(constraint), 2)
11730                         If UBound(LHSOriginalValues(constraint), 1) = UBound(RHSOriginalValues(constraint), 1) Then
11740                             RHS(row) = RHSOriginalValues(constraint)(i, j) - LHSOriginalValues(constraint)(i, j)
11750                         Else
11760                             RHS(row) = RHSOriginalValues(constraint)(j, i) - LHSOriginalValues(constraint)(i, j)
11770                         End If
11780                         row = row + 1
11790                     Next j
11800                 Next i
11810             Else 'if the rhs is anything other than a multicell range
                      'If LHSType(Constraint) = SolverInputType.SingleCellRange Then
                      '    RHS(Row) = RHSOriginalValues(Constraint) - LHSOriginalValues(Constraint)
                      '    Row = Row + 1
                      'Else    'if the left hand side is more than one cell then loop through all lhs cells
11820                 For i = 1 To UBound(LHSOriginalValues(constraint), 1)
11830                     For j = 1 To UBound(LHSOriginalValues(constraint), 2)
11840                         RHS(row) = RHSOriginalValues(constraint) - LHSOriginalValues(constraint)(i, j)
11850                         row = row + 1
11860                     Next j
11870                 Next i
                      ' End If
11880             End If
11890         End If
11900     Next constraint

11910     succeeded = True

Exit Sub:

constraintErrorHandler:
          ' We only trap Escape (Err.Number=18) here; all other errors are passed back to the caller.
          ' Save error message
          Dim ErrorNumber As Long, ErrorDescription As String, ErrorSource As String
10591     ErrorNumber = Err.Number
          If ErrorNumber = 13 Then
              ErrorDescription = Err.Description & " (at constraint " & constraint & " (" & i & ", " & j & "))" & IIf(Erl = 0, "", " (at line " & Erl & ")")
          Else
10601         ErrorDescription = Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
          End If
10611     ErrorSource = Err.Source
10621     If Err.Number = 18 Then
10631         If MsgBox("You have pressed the Escape key. Do you wish to cancel?", _
                         vbCritical + vbYesNo + vbDefaultButton1, _
                         "OpenSolver: User Interrupt Occured...") = vbNo Then
10641             Resume 'continue on from where error occured
10651         Else
                  ' Raise a "user cancelled" error. We cannot use Raise, as that exits immediately without going thru our code below
10661             ErrorNumber = OpenSolver_UserCancelledError
10671             ErrorSource = errorPrefix
10681             ErrorDescription = "Model building cancelled by user."
10691         End If
10701     End If
          
ErrorExit:
          ' Exit, raising an error; none of the following change the Err.Number etc, but we saved them above just in case...
10711     Application.Cursor = xlDefault
10721     Application.StatusBar = False ' Resume normal status bar behaviour
10731     Application.ScreenUpdating = True
          Dim oldCalculationMode As Integer
10741     Application.Calculation = oldCalculationMode
10751     Err.Raise ErrorNumber, ErrorSource, ErrorDescription

      ' Debugging resume; never reached via code
10761     Resume  ' Used for manual trying again after an error

End Sub

Function SolveModel(SolveRelaxation As Boolean) As Integer

          Dim errorPrefix As String
11920     errorPrefix = "OpenSolver Model Solving"

          'Delete any existing nomad log file
          Dim logPath As String
11930     logPath = GetTempFolder & "log1.tmp"
11940     If Dir(logPath) <> "" Then Kill logPath
11950     If Dir(logPath) <> "" Then
11960         On Error GoTo errorHandler
11970         Err.Raise Number:=OpenSolver_SolveError, Source:=errorPrefix, Description:="Unable to delete the log file 'log1.tmp'"
11980     End If
          
          Dim ModelFileName As String
          If Solver = "NeosCBC" Then
11990         ModelFileName = GetModelFileName(True)
          Else
              ModelFileName = GetModelFileName(False)
          End If
          Dim ModelFilePathName As String
          If Solver = "NeosCBC" Then
12000         ModelFilePathName = GetModelFullPath(True)
          Else
              ModelFilePathName = GetModelFullPath(False)
          End If
12010     If Dir(ModelFilePathName) <> "" Then Kill ModelFilePathName ' delete model file, just in case anything goes wrong (and we get left with an old one)
12020     If Dir(ModelFilePathName) <> "" Then
12030         On Error GoTo errorHandler
12040         Err.Raise Number:=OpenSolver_SolveError, Source:=errorPrefix, Description:="Unable to delete file " & ModelFilePathName
12050     End If

          'Delete the CBC solution file if it exists
          Dim SolutionFilePathName As String, SolutionFileName As String
12060     SolutionFileName = GetSolutionFileName
12070     SolutionFilePathName = GetSolutionFullPath
12080     If Dir(SolutionFilePathName) <> "" Then Kill SolutionFilePathName ' delete solution file
12090     If Dir(SolutionFilePathName) <> "" Then
12100         On Error GoTo errorHandler
12110         Err.Raise Number:=OpenSolver_SolveError, Source:=errorPrefix, Description:="Unable to delete the " & SolverString & " solver solution file: " & SolutionFilePathName & ". The problem cannot be solved."
12120     End If

          'Delete the Gurobi solution file if it exists
          Dim GurobiSolutionFilePathName As String, GurobiSolutionFileName As String
12130     GurobiSolutionFileName = Replace(SolutionFileName, ".txt", ".sol")
12140     GurobiSolutionFilePathName = Replace(SolutionFilePathName, ".txt", ".sol")
12150     If Dir(GurobiSolutionFilePathName) <> "" Then Kill GurobiSolutionFilePathName ' delete solution file
12160     If Dir(GurobiSolutionFilePathName) <> "" Then
12170         On Error GoTo errorHandler
12180         Err.Raise Number:=OpenSolver_SolveError, Source:=errorPrefix, Description:="Unable to delete the " & SolverString & " solver solution file: " & GurobiSolutionFilePathName & ". The problem cannot be solved."
12190     End If
          
          'Delete the CBC or Gurobi files with sensitivity data on it if they exist
12200     If Dir(Replace(SolutionFilePathName, "modelsolution", "costranges")) <> "" Then Kill Replace(SolutionFilePathName, "modelsolution", "costranges")
12210     If Dir(Replace(SolutionFilePathName, "modelsolution", "rhsranges")) <> "" Then Kill Replace(SolutionFilePathName, "modelsolution", "rhsranges")
12220     If Dir(Replace(GurobiSolutionFilePathName, "modelsolution", "sensitivityData")) <> "" Then Kill Replace(GurobiSolutionFilePathName, "modelsolution", "sensitivityData")
          'Delete the cbc batch file if it exists
          If Dir(Replace(SolutionFilePathName, "modelsolution.txt", "cbc.bat")) <> "" Then Kill Replace(SolutionFilePathName, "modelsolution.txt", "cbc.bat")

12230     If Solver = "Nomad" Then
12240         SolveRelaxedModel = SolveRelaxation
12250         Set OpenSolverNonLinear.OpenSolver = Me
12260         SolveModel = SolveModel_Nomad(SolveRelaxation)
12270     Else
12280         SolveModel = SolveModel_Linear(SolveRelaxation)
12290     End If
12300     Exit Function
          
errorHandler:
          ' We only trap Escape (Err.Number=18) here; all other errors are passed back to the caller.
          ' Save error message
          Dim ErrorNumber As Long, ErrorDescription As String, ErrorSource As String
12310     ErrorNumber = Err.Number
12320     ErrorDescription = Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
12330     ErrorSource = Err.Source
12340     If Err.Number = 18 Then
12350         If MsgBox("You have pressed the Escape key. Do you wish to cancel?", _
                         vbCritical + vbYesNo + vbDefaultButton1, _
                         "OpenSolver: User Interrupt Occured...") = vbNo Then
12360             Resume 'continue on from where error occured
12370         Else
                  ' Raise a "user cancelled" error. We cannot use Raise, as that exits immediately without going thru our code below
12380             ErrorNumber = OpenSolver_UserCancelledError
12390             ErrorSource = errorPrefix
12400             ErrorDescription = "Model solve cancelled by user."
12410         End If
12420     End If
          
ErrorExit:
          ' Exit, raising an error. None of the following actions change the Err.Number etc, but we saved them above just in case...
12430     Application.Cursor = xlDefault
12440     Application.StatusBar = False ' Resume normal status bar behaviour
12450     Application.ScreenUpdating = True
12460     Err.Raise ErrorNumber, ErrorSource, ErrorDescription

End Function

Function SolveModel_Linear(SolveRelaxation As Boolean) As Integer
12470     SolveStatus = OpenSolverResult.Unsolved
12480     SolveStatusString = "Unsolved"
12490     SolveStatusComment = ""
12500     LinearSolveStatus = LinearSolveResult.Unsolved
12510     LinearSolveStatusString = "Unsolved"
12520     LinearSolutionWasLoaded = False
          
          Dim oldCalculationMode As Integer
12530     oldCalculationMode = Application.Calculation
12540     Application.Calculation = xlCalculationManual
          
          Dim errorPrefix As String
12550     errorPrefix = "OpenSolver Model Solving"
12560     On Error GoTo errorHandler

12570     If ModelStatus <> ModelStatus_Built Then
12580         Err.Raise Number:=OpenSolver_SolveError, Source:=errorPrefix, Description:="The model cannot be solved as it has not yet been built."
12590     End If

          ' Trap Escape key
12600     Application.EnableCancelKey = xlErrorHandler
          
12610     Application.StatusBar = "OpenSolver: Writing Model to disk... " & numVars & " vars, " & NumRows & " rows."
                        
          Dim ExternalSolverPathName As String
          
          If Solver <> "NeosCBC" Then
12620         GetExternalSolverPathName ExternalSolverPathName, Solver ' May throw an error
          End If
          
          'Check that gurobi has been installed
12630     If Solver = "gurobi.bat" Then
12640         If Environ("GUROBI_HOME") = "" Then
12650             Err.Raise Number:=OpenSolver_SolveError, Source:=errorPrefix, Description:="Gurobi does not seem to be installed. Please install from the gurobi website or use the default CBC solver."
12660         End If
12670     End If
          
          Dim i As Long, j As Long, var As Long, row As Long, coeff As Double
          
          ' output the model to an LP format text file
          ' See http://lpsolve.sourceforge.net/5.5/CPLEX-format.htm
          Dim ModelFileName As String
          Dim ModelFilePathName As String
          If Solver = "NeosCBC" Then
12680         ModelFileName = GetModelFileName(True)
12690         ModelFilePathName = GetModelFullPath(True)
          Else
              ModelFileName = GetModelFileName(False)
              ModelFilePathName = GetModelFullPath(False)
          End If

          'This is where we write the .lp or .xml file that has the model that will be solved by our solver
          Dim ErrorString As String
          
          If Solver = "NeosCBC" Then
12701         Call WriteAMPLFile(ModelFilePathName, SolveRelaxation, ErrorString)
              Call CallNEOS(ModelFilePathName, ErrorString)
          Else
12700         Call WriteLPFile(ModelFilePathName, SolveRelaxation, ErrorString)
          
12710         If ErrorString = "ExitSub" Then
12720             GoTo ExitSub
12730         ElseIf ErrorString <> "" Then
12740             On Error GoTo errorHandler
12750             Err.Raise Number:=OpenSolver_SolveError, Source:=errorPrefix, Description:=ErrorString
12760         End If
                  
              ' This is where we solve the problem
              ' Command line for cbc is: cbc model.lp -solution solution.txt
              ' Command line for gurobi is: gurobi_cl ResultFile=solution.sol model.lp
              ' New command line for gurobi is (which runs using our python file and the gurobi interactive shell):
              '                     gurobi.bat gurobiOSRun.py
              '         gurobiOSRun.py is a python file that runs and writes the modelsolution.sol file as well as sensitivityData if they want duals
              
              
              ' Get the Solver Options, stored in named ranges with values such as "=0.12"
              ' Because these are NAMEs, they are always in English, not the local language, so get their value using Val
              Dim SolveOptions As SolveOptionsType
12770         GetSolveOptions sheetName, SolveOptions, ErrorString
12780         If ErrorString <> "" Then
12790             On Error GoTo errorHandler
12800             Err.Raise Number:=OpenSolver_SolveError, Source:=errorPrefix, Description:=ErrorString
12810         End If
              
12820         Application.StatusBar = "OpenSolver: Solving " & IIf(SolveRelaxation, "Relaxed ", "") & "Model... " & numVars & " vars, " & NumIntVars & " int vars " _
                      & "(" & NumBinVars & " 0/1), " & NumRows & " rows, " & SolveOptions.maxTime & "s time limit, " _
                      & SolveOptions.Tolerance * 100 & "% tolerance."
              
              ' We allow some user defined CBC parameters
              Dim CBCExtraParametersString As String
12830         CBCExtraParametersString = GetCBCExtraParametersString(sheet, ErrorString)
12840         If ErrorString <> "" Then
12850             On Error GoTo errorHandler
12860             Err.Raise Number:=OpenSolver_SolveError, Source:=errorPrefix, Description:=ErrorString
12870         End If
                 
              ' If the user has a defined name OpenSolver_ConstraintList, then we need to get the duals by passing an extra option to CBC
              Dim rConstraintList As Range, bGetDuals As Boolean, DualsOnNewSheet As Boolean, DualsOnSameSheet As Boolean
12880         DualsOnSameSheet = GetNamedRangeIfExistsOnSheet(ActiveSheet, "OpenSolver_Duals", rConstraintList)
12890         On Error Resume Next
12900         DualsOnNewSheet = Mid(Names("'" & Replace(ActiveWorkbook.ActiveSheet.Name, "'", "''") & "'!OpenSolver_DualsNewSheet").value, 2) 'frmModel.chkGetDuals2.Value
12910         If Err.Number <> 0 Then
12920             Call SetNameOnSheet("OpenSolver_DualsNewSheet", "=FALSE")
12930             DualsOnNewSheet = False
12940         End If
12950         On Error GoTo errorHandler
12960         bGetDuals = False
12970         If (IntegerCellsRange Is Nothing And BinaryCellsRange Is Nothing) And (DualsOnNewSheet Or DualsOnSameSheet) Then
12980             bGetDuals = True
12990         End If
              
              Dim SolutionFilePathName As String, SolutionFileName As String
13000         SolutionFileName = GetSolutionFileName
13010         SolutionFilePathName = GetSolutionFullPath
                     
              Dim CommandLineRunString As String, PrintingOptionString As String
13020         If Solver = "cbc.exe" Then
                  ' have to split up the command line as excel couldn't have a string longer than 255 characters??
13030             CommandLineRunString = " -directory " & GetTempFolder _
                               & " -import " & ModelFileName _
                               & " -ratioGap " & str(SolveOptions.Tolerance) _
                               & " -seconds " & str(SolveOptions.maxTime) _
                               & CBCExtraParametersString _
                               & " -solve " _
                               & IIf(bGetDuals, " -printingOptions all ", "") _
                               & " -solution " & SolutionFileName
                  '-------------------sensitivity analysis-----------------------------------------------------------
                  'extra command line option of -printingOptions rhs -solution rhsranges.txt gives the allowable increase for constraint rhs.
                  '-this file has the increase as the third input and allowable decrease as the fifth input
                  'extra command line option of -printingOptions objective - solution costranges.txt outputs the ranges on the costs to the costranges file
                  '-this file has the increase as the fifth input and decrease as the third input
13040             PrintingOptionString = IIf(bGetDuals, " -printingOptions rhs  -solution rhsranges.txt -printingOptions objective -solution costranges.txt", "")
                  Open GetTempFolder & "cbc.bat" For Output As 1
                  Print #1, "@echo off" & vbCrLf & """" & ExternalSolverPathName & """" & CommandLineRunString & PrintingOptionString
                  Close #1
                  ExternalSolverPathName = GetTempFolder & "cbc.bat"
                  CommandLineRunString = ""
                  PrintingOptionString = ""
13050         ElseIf Solver = "gurobi.bat" Then
                  Dim CombinedPathName As String
13060             If GetExistingFilePathName(ThisWorkbook.Path, "gurobiOSRun.py", CombinedPathName) Then
13070                 CommandLineRunString = " """ & CombinedPathName & """" ' fix error for Gurobi with spaces in path name by quoting path
                      '========================================================================================
                      'Gurobi can also be run at the command line using gurobi_cl with the following commands
                      '
                      'Solver="gurobi_cl.exe"
                      'CommandLineRunString = " Threads=1" & " TimeLimit=" & Replace(Str(SolveOptions.maxTime), " ", "") & " ResultFile="
                      '             & GetTempFolder & Replace(SolutionFileName, ".txt", ".sol") _
                      '             & " " & ModelFilePathName
                      '
                      '========================================================================================
13080                 PrintingOptionString = ""
13090             Else
13100                 Err.Raise Number:=OpenSolver_SolveError, Source:=errorPrefix, Description:="Could not find the file 'GurobiOSRun.py' in the folder with 'OpenSolver.xlam'. This file should have come with your OpenSolver download. Please move it into the OpenSolver folder or download it from the Opensolver website."
13110             End If
13120         ElseIf Solver = "cplex.exe" Then
13130             CommandLineRunString = """" & GetTempFolder & ModelFilePathName & """" & " write=" & Replace(SolutionFileName, ".txt", ".sol")
13140             PrintingOptionString = ""
13150         End If
              
              Dim logCommand As String, logFileName As String
              logFileName = "log1.tmp"
              logCommand = " > " & """" & GetTempFolder & logFileName & """"
                  
              Dim ExecutionCompleted As Boolean
              'ExecutionCompleted = OSSolveSync(CBCRunString, IIf(ShowIterationResults, SW_SHOWNORMAL, SW_SHOWMINIMIZED), True) ' This stops escape being detected properly unless Excel is first 'activated?' with a click
13160         ExternalSolverPathName = """" & ExternalSolverPathName & """" ' fix error for Gurobi with spaces in path name by quoting path
              Dim exeResult As Long, userCancelled As Boolean
13170         ExecutionCompleted = OSSolveSync(ExternalSolverPathName, CommandLineRunString, PrintingOptionString, logCommand, IIf(SolveOptions.ShowIterationResults, SW_SHOWNORMAL, SW_HIDE), True, userCancelled, exeResult) ' Run CBC, waiting for completion
13180         If userCancelled Then
                  ' User pressed escape. Dialogs have already been shown. Exit with a 'cancelled' error
13190             On Error GoTo errorHandler
13200             Err.Raise Number:=OpenSolver_UserCancelledError, Source:=errorPrefix, Description:="The solving process was cancelled by the user."
13210         End If
13220         If exeResult <> 0 Then
                  ' User pressed escape. Dialogs have already been shown. Exit with a 'cancelled' error
13230             On Error GoTo errorHandler
13240             Err.Raise Number:=OpenSolver_SolveError, Source:=errorPrefix, Description:="The " & SolverString & " solver did not complete, but aborted with the error code " & exeResult & "." & vbCrLf & vbCrLf & "The last log file can be viewed under the OpenSolver menu and may give you more information on what caused this error."
13250         End If
              
             ' Application.StatusBar = "OpenSolver: Loading Solution... " & NumVars & " vars, " & NumRows & " rows."
              
              ' Read in the solution, status first
              ' Statuses can be some combination of (based on CBC.exe text strings):
              '    Optimal
              '    Stopped on time
              '    Stopped on iterations
              '    Stopped on difficulties
              '    Stopped on ctrl-c
              '     (no integer solution - continuous used)
            
              'Read the solution files and get back answers as well as dual information if it is wanted
13260         ErrorString = ""
13270         If Solver = "gurobi.bat" Then
                  '=============Read the gurobi model file==========================
13280             SolutionFilePathName = Replace(SolutionFilePathName, ".txt", ".sol")
13290             If Dir(SolutionFilePathName) = "" Then
13300                 On Error GoTo errorHandler
13310                 Err.Raise Number:=OpenSolver_SolveError, Source:=errorPrefix, Description:="The Gurobi solver did not create a solution file. No new solution is available."
13320                 GoTo ExitSub
13330             End If
                  Dim VariableValues() As Double, solutionLoaded As Boolean
13340             solutionLoaded = ReadGurobiModel(SolutionFilePathName, bGetDuals, ErrorString)
13350             On Error GoTo errorHandler
13360             If ErrorString <> "" Then
13370                 Err.Raise Number:=OpenSolver_SolveError, Source:=errorPrefix, Description:=ErrorString
13380             ElseIf Not solutionLoaded Then 'read error
13390                 GoTo ExitSub
          '                  Err.Raise Number:=OpenSolver_SolveError, Source:=errorPrefix, Description:="OpenSolver encountered an error while trying to read the gurobi solution file at " & SolutionFilePathName & ". No solution available on the spreadsheet."
13400             End If
13410         ElseIf Solver = "cbc.exe" Then
                  '===============Read the CBC model file============================
13420             If Dir(SolutionFilePathName) = "" Then
13430                 On Error GoTo errorHandler
13440                 Err.Raise Number:=OpenSolver_SolveError, Source:=errorPrefix, Description:="The CBC solver did not create a solution file. No new solution is available."
13450                 GoTo ExitSub
13460             End If
                  Dim CBCSolutionLoaded As Boolean
13470             CBCSolutionLoaded = ReadCBCModel(SolutionFilePathName, CommandLineRunString, bGetDuals, _
                                    rConstraintList, ErrorString)
    
13480             If ErrorString <> "" Then
13490                 Err.Raise Number:=OpenSolver_SolveError, Source:=errorPrefix, Description:=ErrorString
13500             ElseIf Not CBCSolutionLoaded Then
13510                 GoTo ExitSub
13520             End If
13530         ElseIf Solver = "cplex.exe" Then
13540             SolutionFilePathName = Replace(SolutionFilePathName, ".txt", ".sol")
13550             If Dir(SolutionFilePathName) = "" Then
13560                 On Error GoTo errorHandler
13570                 Err.Raise Number:=OpenSolver_SolveError, Source:=errorPrefix, Description:="The CPLEX solver did not create a solution file. No new solution is available."
13580                 GoTo ExitSub
13590             End If
13600             Err.Raise Number:=OpenSolver_SolveError, Source:=errorPrefix, Description:="The CPLEX solver has not been fully incorporated into OpenSolver yet, you can view the solution in modelsolution.txt in the temporary folder. To view on the sheet please pick abother solver. Sorry for the inconvenience we are working out this problem now."
13610             GoTo ExitSub
13620         End If
          End If
13630     LinearSolutionWasLoaded = True
          
          ' Perform a linearity check unless the user has requested otherwise
          Dim performLinearityCheck As Boolean, s As String
13640     performLinearityCheck = True
13650     If GetNameValueIfExists(ActiveWorkbook, "'" & Replace(ActiveSheet.Name, "'", "''") & "'!OpenSolver_LinearityCheck", s) Then
13660         performLinearityCheck = s = "1"
13670     End If
13680     If LinearSolutionWasLoaded And performLinearityCheck Then
              Dim fullLinearityCheckWasPerformed As Boolean
13690         QuickLinearityCheck fullLinearityCheckWasPerformed
13700         If fullLinearityCheckWasPerformed Then
13710             SolveStatus = OpenSolverResult.AbortedThruUserAction
13720             SolveStatusString = "No Solution Found "
13730             LinearSolveStatus = LinearSolveResult.SolveStopped
13740         End If
13750     End If
            
          '=====================Write Sensitivity Tables=========================================================
13760     If bGetDuals And SolveStatusString = "Optimal" Then
              'write the duals on the same sheet if the user has picked this option
13770         If DualsOnSameSheet Then WriteConstraintListToSheet rConstraintList
              'If the user wants a new sheet with the sensitivity data then call the functions that write this
13780         If DualsOnNewSheet Then
                  Dim nameSheet As String
13790             nameSheet = MakeNewSheet(RawSheetName & " Sensitivity", sheetName)
13800             sheet.Select
13810             Call WriteConstraintSensitivityTable(nameSheet, SolverString)
                  'If you want to end up on the sensitivity page use this next line
                  'Sheets(nameSheet).Select
13820         End If
13830     ElseIf Not bGetDuals And (DualsOnNewSheet Or DualsOnSameSheet) Then
13840         Err.Raise Number:=OpenSolver_SolveError, Source:=errorPrefix, Description:="Could not get sensitivity analysis due to binary and/or integer constraints." & vbCrLf & vbCrLf & "Turn off sensitivity in the model dialogue or reformulate your model without these constraints." & vbCrLf & vbCrLf & "The " & SolverString & " solution has been returned to the sheet." & vbCrLf
13850     End If
          '=========================================================================================================
        

ExitSub:
          ' We can fall thru to here, or jump here if the problem is shown to be infeasible before we run the CBC Solver
13860     Application.Cursor = xlDefault
13870     Application.StatusBar = False ' Resume normal status bar behaviour
13880     Application.ScreenUpdating = True
13890     Application.Calculation = oldCalculationMode
13900     Application.Calculate
13910     Close #1 ' Close any open file; this does not seem to ever give errors
          
13920     SolveModel_Linear = SolveStatus    ' Return the main result
13930     Exit Function

errorHandler:
          ' We only trap Escape (Err.Number=18) here; all other errors are passed back to the caller.
          ' Save error message
          Dim ErrorNumber As Long, ErrorDescription As String, ErrorSource As String
13940     ErrorNumber = Err.Number
13950     ErrorDescription = Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
13960     ErrorSource = Err.Source
13970     If Err.Number = 18 Then
13980         If MsgBox("You have pressed the Escape key. Do you wish to cancel?", _
                         vbCritical + vbYesNo + vbDefaultButton1, _
                         "OpenSolver: User Interrupt Occured...") = vbNo Then
13990             Resume 'continue on from where error occured
14000         Else
                  ' Raise a "user cancelled" error. We cannot use Raise, as that exits immediately without going thru our code below
14010             ErrorNumber = OpenSolver_UserCancelledError
14020             ErrorSource = errorPrefix
14030             ErrorDescription = "Model solve cancelled by user."
14040         End If
14050     End If
          
ErrorExit:
          ' Exit, raising an error. None of the following actions change the Err.Number etc, but we saved them above just in case...
14060     Application.Cursor = xlDefault
14070     Application.StatusBar = False ' Resume normal status bar behaviour
14080     Application.ScreenUpdating = True
14090     Application.Calculation = oldCalculationMode
14100     Application.Calculate
14110     Close #1 ' Close any open file; this does not seem to ever give errors
14120     Err.Raise ErrorNumber, ErrorSource, ErrorDescription

      ' Debugging resume; never reached via code
14130     Resume  ' Used for manual trying again after an error
End Function

Sub WriteLPFile(ModelFilePathName As String, SolveRelaxation As Boolean, ErrorString As String)

          Dim i As Long, j As Long, var As Long, row As Long, coeff As Double, Line As String, Line2 As String, Line3 As String
          
14140     ErrorString = ""
          Dim commentStart As String  'Character for starting comments for chosen solver
14150     commentStart = "\"
          
          ' output the model to an LP format text file
          ' See http://lpsolve.sourceforge.net/5.5/CPLEX-format.htm

14160     Open ModelFilePathName For Output As 1 ' supply path with filename
14170     Print #1, commentStart & " Model solved using the solver '" & Solver & "'"
14180     Print #1, commentStart & " Model for sheet " & left(sheetName, Len(sheetName) - 1) ' remove trailing !
          ' Warning: TimeTakenToBuildModel may be negative if we cross midnight! This is unlikely for short build times
14190     Print #1, commentStart & " It took " & IIf(SecondsTakenToBuildModel > 100, SecondsTakenToBuildModel, TimeTakenToBuildModel) & " seconds to build the model."
14200     Print #1, commentStart & " Model has " & NumConstraints & " Excel constraints giving " & NumRows & " constraint rows and " & numVars & " variables."
14210     If SolveRelaxation And (NumBinVars > 0) Then
14220         Print #1, commentStart & " (Formulation for relaxed problem)"
14230     End If
14240     Print #1, IIf(ObjectiveSense = MaximiseObjective, "MAXIMIZE", "MINIMIZE")   ' We use Minimise for both minimisation, and also for seeking a target (TargetObjective)
14250     Print #1, "Obj:";
14260     If ObjectiveSense = TargetObjective Then
              ' We want the objective to achieve some target value; we have no objective; nothing is output
14270         Print #1, ' A new line meaning a blank objective; add a comment to this effect next
14280         If Solver <> "gurobi.bat" Then
14290             Print #1, commentStart & " We have no objective function as the objective must achieve a given target value"
14300         End If
14310     Else
14320         For var = 1 To numVars
14330             If Abs(CostCoeffs(var)) > ZERO Then Print #1, " " & StrEx(CostCoeffs(var)) & " " & ValidLPFileVarName(VarNames(var));
                  'If Abs(CostCoeffs(var)) > ZERO Then Print #1, " " & IIf(CostCoeffs(var) > 0, "+", "") & CostCoeffs(var) & " " & ValidLPFileVarName(VarNamesCollection(var).AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False));
14340         Next var
14350         Print #1,   ' New line
14360     End If
14370     Print #1,   ' New line
14380     Print #1, "SUBJECT TO"
          
          ' If we are seeking a specific objective value, we add this as a constraint
14390     If ObjectiveSense = TargetObjective Then
14400         Print #1, commentStart & " The objective must achieve a given target value; this constraint enforces this."
              Dim NonTrivialObjective As Boolean
14410         For var = 1 To numVars
14420             If Abs(CostCoeffs(var)) > ZERO Then
14430                 Print #1, " " & StrEx(CostCoeffs(var)) & " " & ValidLPFileVarName(VarNames(var));
14440                 NonTrivialObjective = True
14450             End If
14460         Next var
14470         Print #1, " = " & StrEx(ObjectiveTargetValue)
14480         If Not NonTrivialObjective And ObjectiveTargetValue <> 0 Then
14490             SolveStatus = OpenSolverResult.Infeasible
14500             SolveStatusString = "Infeasible Objective Target"
14510             SolveStatusComment = "The model's objective cell does not depend on the decision variables" _
                        & " and so cannot be adjusted to achieve the target value" & ObjectiveTargetValue & "."
14520             GoTo ExitSub
14530         End If
14540     End If
          
          Dim constraint As Long
          Dim instance As Long
14550     constraint = 1
          
14560     For row = 1 To NumRows
14570         GetConstraintFromRow row, constraint, instance  ' Which Excel constraint are we in, and which instance?
14580         If instance = 1 Then
                  ' We are outputting the first row of a new Excel constraint block; put a comment in the .lp file
14590             Print #1, commentStart & " " & ConstraintSummary(constraint)
14600         End If
         
              'For var = 1 To NumVars
              '    If Abs(a(row, var)) > ZERO Then
              '        Print #1, " " & IIf(a(row, var) > 0, "+", "") & a(row, var) & " " & ValidLPFileVarName(VarNames(var));
              '        ValidRow = True
              '    End If
              'Next var
14610         With SparseA(row)
14620             For i = 1 To .Count
14630                 var = .Index(i)
14640                 coeff = .Coefficient(i)
14650                 Print #1, " " & StrEx(coeff) & " " & ValidLPFileVarName(VarNames(var));
14660             Next i
14670         End With
              
14680         If SparseA(row).Count = 0 Then
                  'We have a constraint that does not vary with the decision variables; check it is satisifed
                  Dim LHSValue As Double, RHSValue As Double
14690             GetArrayPosnFromConstraintInstance constraint, instance, i, j
14700             LHSValue = LHSOriginalValues(constraint)(i, j)
14710             If RHSType(constraint) <> MultiCellRange Then
14720                 RHSValue = RHSOriginalValues(constraint)
14730             Else
14740                 RHSValue = RHSOriginalValues(constraint)(i, j)
14750             End If
                  ' GetExcelConstraintInfoFromRow Constraint, Row, LHSValue, RHSValue, Probleminstance, sLHS, sRHS
14760             If (Relation(row) = RelationGE And LHSValue < RHSValue - EPSILON) _
                     Or (Relation(row) = RelationEQ And Abs(LHSValue - RHSValue) > EPSILON) _
                     Or (Relation(row) = RelationLE And LHSValue > RHSValue + EPSILON) Then
      '                  ' AJM 28.11.2010 We cannot use any IIF() that references LHSRange(row).Address because all arguments are always evaluated, and LHSRange(row).Address is not defined and throws an error
      '                  sLHS = LHSRange(j).Address & " = " & lhsvalue
                        Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
14770                   GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
                        ' If the RHS is a range, we show its address; if not, RHSString contains the RHS's constant or formula
14780                   If Not RHSCellRange Is Nothing Then RHSstring = RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
14790                   SolveStatus = OpenSolverResult.Infeasible
14800                   SolveStatusString = "Infeasible"
14810                   SolveStatusComment = "The model contains an Excel constraint '" & ConstraintSummary(constraint) & "' for which instance " _
                               & instance & " does not depend on the decision variables and is not satisfied." & vbCrLf _
                              & "Constraint specifies: " _
                               & "LHS=" & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "=" & LHSValue _
                              & " " & SolverRelationAsString(Relation(row)) & " " _
                              & "RHS=" & RHSstring & "=" & RHSValue
14820                   Close #1
14830                   GoTo ExitSub
14840             End If
                   'We output the row as a comment
14850             Print #1, commentStart & " (A row with all zero coeffs)";
14860         End If
14870         If Relation(row) = RelationGE Then
14880             Print #1, " >= ";
14890         ElseIf Relation(row) = RelationLE Then
14900             Print #1, " <= ";
14910         ElseIf Relation(row) = RelationEQ Then
14920             Print #1, " = ";
14930         End If
14940         Print #1, StrEx(RHS(row))
14950     Next row
14960     Print #1,   ' New line
          Dim c As Range
14970     If SolveRelaxation And Not (BinaryCellsRange Is Nothing) Then
14980         Print #1, commentStart & " (Upper bounds of 1 on the relaxed binary variables)"
14990         For Each c In BinaryCellsRange
15000             Print #1, ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False)); " <= 1"
15010         Next c
15020         Print #1,   ' New line
15030     End If
          
          ' Output the bounds; this should happen before we output the GENERAL or INTEGER sections (at least for CPLEX .lp files)
          ' See http://lpsolve.sourceforge.net/
          ' The LP file assumes lower bounds on all variables unless we tell it otherwise.
15040     If Not AssumeNonNegativeVars Then
              ' We need to make all variables FREE variables (i.e. no lower bounds), except for the Binary variables
15050         Print #1, "BOUNDS"
15060         Print #1, commentStart & "'Assume Non Negative' is FALSE, so default lower bounds of zero are removed from all non-binary variables."
15070         If BinaryCellsRange Is Nothing Then
15080             For Each c In AdjustableCells
15090                 Print #1, " " & ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False)) & " FREE"
15100             Next c
15110         Else
15120             For Each c In AdjustableCells
15130                 If Intersect(BinaryCellsRange, c) Is Nothing Then
15140                     Print #1, " " & ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False)) & " FREE"
15150                 End If
15160             Next c
15170         End If
15180         Print #1,   ' New line
15190     Else
              ' If AssumeNonNegative, then we need to apply lower bounds to any variables without explicit lower bounds.
              ' Excel 2007 requires the following conditions to be met for a user-supplied lower bound to exist
              '  The LHS range includes a decision variable directly (not via a formula referring to the decision var).
              '  The RHS is either: a constant, or: a (single or multicell) range that containing a constant or a formula. The RHS may not
              '  be specified as a formula (even one evaluating to a constant)
              '  The resulting equation does indeed just depend on a single decision variable (ie the RHS range turned out to be constant)
              ' Note: Excel 2010 seems a bit different in that the WHOLE LHS range must be decision variables; this needs to be confirmed
              ' Note: A constraint of the form 'range <= decision var' does NOT count as a lower bound; the decision var MUST be in the LHS
              Dim BoundedVariables As Range
15200         For row = 1 To NumRows
15210             If Relation(row) = RelationGE And SparseA(row).Count = 1 Then
                      ' This is a constraint apply a lower bound to one variable. Check that the LHS did directly specify the decision variable.
15220                 var = SparseA(row).Index(1)
15230                 GetConstraintFromRow row, constraint, instance
15240                 If RHSType(constraint) <> Formula Then
15250                     GetArrayPosnFromConstraintInstance constraint, instance, i, j
15260                     If LHSRange(constraint)(i, j).Address(RowAbsolute:=False, ColumnAbsolute:=False) = VarNames(var) _
                             And LHSRange(constraint).Worksheet.Name = AdjustableCells.Worksheet.Name Then ' AJM Must be on the same sheet
                              ' This is a direct lower bound on this variable
15270                         Set BoundedVariables = ProperUnion(BoundedVariables, LHSRange(constraint)(i, j))
15280                     End If
15290                 End If
15300             End If
15310         Next row
              ' We need to mark variables with explicit lower bounds as FREE variables (allowing the possibly negative lower bound to be effective).
              ' However, we don't make Binary variables free
15320         If Not BoundedVariables Is Nothing Then
15330             Print #1, "BOUNDS"
15340             Print #1, commentStart & "'Assume Non Negative' is TRUE, so default lower bounds of zero are removed only from non-binary variables already given explicit lower bounds."
15350             If BinaryCellsRange Is Nothing Then
15360                 For Each c In BoundedVariables
15370                     Print #1, " " & ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False)) & " FREE"
15380                 Next c
15390             Else
15400                 For Each c In BoundedVariables
15410                     If Intersect(BinaryCellsRange, c) Is Nothing Then
15420                         Print #1, " " & ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False)) & " FREE"
15430                     End If
15440                 Next c
15450             End If
15460             Print #1,   ' New line
15470         End If
15480     End If
          
          ' Output any integer variables. If the binary and integer variables overlap, then we check each integer variable to ensure it is not binary
15490     If Not SolveRelaxation And Not (IntegerCellsRange Is Nothing) Then
15500         If Not BinaryCellsRange Is Nothing Then
15510             On Error Resume Next
15520             If Not Intersect(BinaryCellsRange, IntegerCellsRange).Count = IntegerCellsRange.Count Then
15530                 If Err.Number = 91 Then
                          'If there is no intersect Excel throw an error
15540                     Print #1, "GENERAL"
15550                     For Each c In IntegerCellsRange
15560                             Print #1, " " & ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False));
15570                     Next c
15580                     On Error GoTo errorHandler
15590                 Else    'if there is an intersect between the binary cell ranges and integer cell ranges
15600                     On Error GoTo errorHandler
15610                     Print #1, "GENERAL"
15620                     For Each c In IntegerCellsRange
15630                         If Intersect(c, BinaryCellsRange) Is Nothing Then
15640                             Print #1, " " & ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False));
15650                         End If
15660                     Next c
15670                 End If
15680             End If
15690          Else   'if there no binary cell range
15700             Print #1, "GENERAL"
15710             For Each c In IntegerCellsRange
15720                     Print #1, " " & ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False));
15730             Next c
15740         End If
15750         Print #1, ' New line
15760     End If
          
          ' Output binary variables
15770     If Not SolveRelaxation And Not (BinaryCellsRange Is Nothing) Then
15780         Print #1, "BINARY"
15790         For Each c In BinaryCellsRange
15800             Print #1, " " & ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False));
15810         Next c
15820         Print #1,   ' New line
15830     End If
15840     Print #1,   ' New line
          
15850     Print #1, "END"
          
15860     Close 1

15870     Exit Sub
          
ExitSub:
15880     ErrorString = "ExitSub"
15890     Exit Sub
errorHandler:
15900     ErrorString = "Error writing LP File."

End Sub

Sub WriteAMPLFile(ModelFilePathName As String, SolveRelaxation As Boolean, ErrorString As String)

     Dim i As Long, j As Long, var As Long, row As Long, coeff As Double, c As Range, Line As String, Line1 As String, Line2 As String, Line3 As String
     Dim VarDic As Collection
     
     Set VarDic = New Collection

     ErrorString = ""
     Dim commentStart As String  'Character for starting comments for chosen solver
     commentStart = "#"

     Open ModelFilePathName For Output As 1 ' supply path with filename
     
     ' XML
     Print #1, "&lt;document&gt;"
     Print #1, "&lt;category&gt;milp&lt;/category&gt;"
     Print #1, "&lt;solver&gt;cbc&lt;/solver&gt;"
     Print #1, "&lt;inputType&gt;AMPL&lt;/inputType&gt;"
     Print #1, "&lt;client&gt;&lt;/client&gt;"
     Print #1, "&lt;priority&gt;short&lt;/priority&gt;"
     Print #1, "&lt;email&gt;&lt;/email&gt;"
     
     ' Model File - Replace with Data File
     Print #1, "&lt;model&gt;&lt;![CDATA[# Define our sets, parameters and variables (with names matching those"
     Print #1, "# used in defining the data items)"
     
     ' Variables
     For var = 1 To numVars
         VarDic.Add "", ValidLPFileVarName(VarNames(var))
     Next var
     
     ' Sets - Vars
     If Not SolveRelaxation Then
         If Not IntegerCellsRange Is Nothing Then
             For Each c In IntegerCellsRange
                 VarDic.Remove (ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False)))
                 VarDic.Add ", integer", ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False))
             Next c
         End If
          
         Line3 = ""
         If Not BinaryCellsRange Is Nothing Then
             For Each c In BinaryCellsRange
                 VarDic.Remove (ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False)))
                 VarDic.Add ", binary", ValidLPFileVarName(c.Address(RowAbsolute:=False, ColumnAbsolute:=False))
             Next c
         End If
     End If
     
     ' Intialise Variables
     For var = 1 To numVars
         Print #1, "var " & ValidLPFileVarName(VarNames(var)) & VarDic.Item(ValidLPFileVarName(VarNames(var))) & ";"
     Next var
     
     ' Parameter - Costs
     Print #1,   ' New line
     Line1 = "  "
     For var = 1 To numVars
         Line = Line & ValidLPFileVarName(VarNames(var)) & "*" & CostCoeffs(var)
         If var < numVars Then
             Line = Line & " + "
         End If
     Next var
     
     ' Objective function replaced with constraint if
     If ObjectiveSense = TargetObjective Then
         Print #1, commentStart & " We have no objective function as the objective must achieve a given target value"
         Print #1,
         
         Print #1, commentStart & " The objective must achieve a given target value; this constraint enforces this."
         Print #1, "subject to TargetConstr {c in Constraints}:"
         Print #1, Line * " = " & StrEx(ObjectiveTargetValue) & ";"
     Else
         ' Determine objective direction
         If ObjectiveSense = MaximiseObjective Then
            Print #1, "maximize Total_Cost:"
         Else
            Print #1, "minimize Total_Cost:"
         End If
         
         Print #1, Line & ";"
         Print #1,   ' New line
     End If
     
     ' Subject to Constraints
     Dim constraint As Long
     Dim instance As Long
     constraint = 1
     
     ' Loop through constraints
     For row = 1 To NumRows
         Line = "   "
         GetConstraintFromRow row, constraint, instance  ' Which Excel constraint are we in, and which instance?
         If instance = 1 Then
             ' We are outputting the first row of a new Excel constraint block; put a comment in the .lp file
             Print #1, commentStart & " " & Replace(Replace(ConstraintSummary(constraint), "<", "&lt;"), ">", "&gt;")
         End If
         
         ' Gather variables
         With SparseA(row)
             For i = 1 To .Count
                 var = .Index(i)
                 coeff = .Coefficient(i)
                 Line = Line & StrEx(coeff) & " * " & ValidLPFileVarName(VarNames(var)) & " "
             Next i
         End With
         
         ' Ensure there were some non-zero coeffs
         If Line <> "   " Then
            
            ' Print constraint header
            Print #1, "subject to c" & row & ":"
            
            ' Check sense
            If Relation(row) = RelationGE Then
                Line = Line & "&gt;= "
            ElseIf Relation(row) = RelationLE Then
                Line = Line & "&lt;= "
            ElseIf Relation(row) = RelationEQ Then
                Line = Line & "== "
            End If
            
            ' Print with RHS
            Print #1, Line & StrEx(RHS(row))
            
            If SparseA(row).Count = 0 Then
                'We have a constraint that does not vary with the decision variables; check it is satisifed
                Dim LHSValue As Double, RHSValue As Double
                GetArrayPosnFromConstraintInstance constraint, instance, i, j
                LHSValue = LHSOriginalValues(constraint)(i, j)
                If RHSType(constraint) <> MultiCellRange Then
                    RHSValue = RHSOriginalValues(constraint)
                Else
                    RHSValue = RHSOriginalValues(constraint)(i, j)
                End If
                ' GetExcelConstraintInfoFromRow Constraint, Row, LHSValue, RHSValue, Probleminstance, sLHS, sRHS
                If (Relation(row) = RelationGE And LHSValue < RHSValue - EPSILON) _
                   Or (Relation(row) = RelationEQ And Abs(LHSValue - RHSValue) > EPSILON) _
                   Or (Relation(row) = RelationLE And LHSValue > RHSValue + EPSILON) Then
    '                  ' AJM 28.11.2010 We cannot use any IIF() that references LHSRange(row).Address because all arguments are always evaluated, and LHSRange(row).Address is not defined and throws an error
    '                  sLHS = LHSRange(j).Address & " = " & lhsvalue
                      Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
                      GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
                      ' If the RHS is a range, we show its address; if not, RHSString contains the RHS's constant or formula
                      If Not RHSCellRange Is Nothing Then RHSstring = RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
                      SolveStatus = OpenSolverResult.Infeasible
                      SolveStatusString = "Infeasible"
                      SolveStatusComment = "The model contains an Excel constraint '" & ConstraintSummary(constraint) & "' for which instance " _
                             & instance & " does not depend on the decision variables and is not satisfied." & vbCrLf _
                            & "Constraint specifies: " _
                             & "LHS=" & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "=" & LHSValue _
                            & " " & SolverRelationAsString(Relation(row)) & " " _
                            & "RHS=" & RHSstring & "=" & RHSValue
                      Close #1
                      GoTo ExitSub
                End If
                 'We output the row as a comment
                Print #1, commentStart & " (A row with all zero coeffs)";
            End If
        
            Print #1, ";"
           
            Print #1,   ' New line
         End If
     Next row
     
     ' Run Commands
     Print #1,   ' New line
     Print #1, commentStart & " Solve the problem"
     Print #1, "option solver cbc;"
     Print #1, "solve;"
     Print #1,   ' New line
     ' Display variables
     For var = 1 To numVars
         Print #1, "_display " & ValidLPFileVarName(VarNames(var)) & ";"
     Next var
     Print #1, "_display solve_result;"
     Print #1,   ' New line
     Print #1, "end]]&gt;&lt;/model&gt;"
     
     ' Closing XML
     Print #1, "&lt;data&gt;&lt;![CDATA[]]&gt;&lt;/data&gt;"
     Print #1,   ' New line
     Print #1, "&lt;commands&gt;&lt;![CDATA[]]&gt;&lt;/commands&gt;"
     Print #1,   ' New line
     Print #1, "&lt;comments&gt;&lt;![CDATA[]]&gt;&lt;/comments&gt;"
     Print #1,   ' New line
     Print #1, "&lt;/document&gt;"
     
     Close 1

     Exit Sub
          
ExitSub:
     ErrorString = "ExitSub"
     Exit Sub
errorHandler:
     ErrorString = "Error writing data File."

End Sub

Sub CallNEOS(ModelFilePathName As String, ErrorString As String)
     Dim objSvrHTTP As MSXML2.ServerXMLHTTP60, message As String, txtURL As String
     Dim Done As Boolean, result As String
     Dim openingParen As String, closingParen As String, jobNumber As String, Password As String, solutionFile As String, solution As String
     Dim i As Integer, LinearSolveStatusString As String, var As Long
     
     ' Server name
     txtURL = "http://www.neos-server.org:3332"
     Set objSvrHTTP = New MSXML2.ServerXMLHTTP60
     
     ' Set up obj for a POST request
     objSvrHTTP.Open "POST", txtURL, False
     
     ' Import file as continuous string
     Open ModelFilePathName For Input As #1
         message = Input$(LOF(1), 1)
     Close #1
     
     ' Set up message as XML
     message = "<methodCall><methodName>submitJob</methodName><params><param><value><string>" _
        & message & "</string></value></param></params></methodCall>"
     
     ' Send Message to NEOS
     objSvrHTTP.send message
     
     ' Extract Job Number
     openingParen = InStr(objSvrHTTP.responseText, "<int>")
     closingParen = InStr(objSvrHTTP.responseText, "</int>")
     jobNumber = Mid(objSvrHTTP.responseText, openingParen + Len("<int>"), closingParen - openingParen - Len("<int>"))
     
     If jobNumber = 0 Then
         MsgBox "An error occured when sending file to NEOS."
         GoTo ExitSub
     End If
     
     ' Extract Password
     openingParen = InStr(objSvrHTTP.responseText, "<string>")
     closingParen = InStr(objSvrHTTP.responseText, "</string>")
     Password = Mid(objSvrHTTP.responseText, openingParen + Len("<string>"), closingParen - openingParen - Len("<string>"))
     
     ' Set up Job Status message for XML
     message = "<methodCall><methodName>getJobStatus</methodName><params><param><value><int>" _
        & jobNumber & "</int></value><value><string>" & Password & _
        "</string></value></param></params></methodCall>"
     Done = False
     
     CallingNeos.Show False
     
     ' Loop until job is done
     While Done = False
         DoEvents
         
         ' Reset obj
         Set objSvrHTTP = New MSXML2.ServerXMLHTTP60
         objSvrHTTP.Open "POST", txtURL, False
         
         ' Send message
         objSvrHTTP.send message
         
         ' Extract answer
         openingParen = InStr(objSvrHTTP.responseText, "<string>")
         closingParen = InStr(objSvrHTTP.responseText, "</string>")
         result = Mid(objSvrHTTP.responseText, openingParen + 8, closingParen - openingParen - 8)
         
         ' Evaluate result
         If result = "Done" Then
             Done = True
         ElseIf result <> "Waiting" And result <> "Running" Then
             MsgBox "An error occured when sending file to NEOS. Neos returned: " & result
             GoTo ExitSub
         Else
             Application.Wait (Now + TimeValue("0:00:01"))
         End If
     Wend
     
     CallingNeos.Hide
     
     ' Set up final message for XML
     message = "<methodCall><methodName>getFinalResults</methodName><params><param><value><int>" _
        & jobNumber & "</int></value></param><param><value><string>" & Password & _
        "</string></value></param></params></methodCall>"
     
     ' Reset obj
     Set objSvrHTTP = New MSXML2.ServerXMLHTTP60
     objSvrHTTP.Open "POST", txtURL, False
    
     objSvrHTTP.send message
     
     ' Extract Result
     openingParen = InStr(objSvrHTTP.responseText, "<base64>")
     closingParen = InStr(objSvrHTTP.responseText, "</base64>")
     result = Mid(objSvrHTTP.responseText, openingParen + 8, closingParen - openingParen - 8)
     
     ' The message is returned from NEOS in base 64
     solution = DecodeBase64(result)
    
     ' Extract the solve status
     openingParen = InStr(solution, "solve_result")
     LinearSolveStatusString = right(solution, Len(solution) - openingParen - Len("solve_result"))
    
     ' Determine Feasibility
     If LinearSolveStatusString Like "*unbounded*" Then
         SolveStatus = OpenSolverResult.Unbounded
         SolveStatusString = "No Solution Found (Unbounded)"
         LinearSolveStatus = LinearSolveResult.Unbounded
         '
     ElseIf LinearSolveStatusString Like "*infeasible*" Then ' Stopped on iterations or time
         SolveStatus = OpenSolverResult.Infeasible
         SolveStatusString = "No Feasible Solution"
         LinearSolveStatus = LinearSolveResult.Infeasible
         GoTo ExitSub
         '
     ElseIf Not LinearSolveStatusString Like "*solved*" Then
         openingParen = InStr(solution, ">>>")
         If openingParen = 0 Then
             openingParen = InStr(solution, "processing commands.")
             LinearSolveStatusString = right(solution, Len(solution) - openingParen - Len("processing commands."))
         Else
             closingParen = InStr(solution, "<<<")
             LinearSolveStatusString = "Error: " & Mid(solution, openingParen, closingParen - openingParen)
         End If
         GoTo ExitSub
     Else
         SolveStatus = OpenSolverResult.Optimal
         SolveStatusString = "Optimal"
         LinearSolveStatus = LinearSolveResult.Optimal
         '
     End If
     
     Application.Calculation = xlCalculationAutomatic
    
     ' Display results to sheet
     For var = 1 To numVars
         openingParen = InStr(solution, ValidLPFileVarName(VarNames(var)))
         closingParen = openingParen + InStr(right(solution, Len(solution) - openingParen), "_display")
         result = Mid(solution, openingParen + Len(ValidLPFileVarName(VarNames(var))) + 1, Application.Max(closingParen - openingParen - Len(ValidLPFileVarName(VarNames(var))) - 1, 0))
         
         ' Converting result to number
         AdjustableCells.Worksheet.Range(VarNames(var)).Value2 = "= " & result & " * 1"
         
         ' Removing equal sign
         AdjustableCells.Worksheet.Range(VarNames(var)).Value2 = AdjustableCells.Worksheet.Range(VarNames(var)).Value2
     Next
     
     Application.Calculation = xlCalculationManual
     
15870          Exit Sub
          
ExitSub:
15880     ErrorString = "ExitSub"
15890     Exit Sub
errorHandler:
15900     ErrorString = "Error contacting NEOS."

End Sub

' Code by Tim Hastings
Private Function DecodeBase64(ByVal strData As String) As String
    Dim objXML As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMElement
  
    ' Help from MSXML
    Set objXML = New MSXML2.DOMDocument60
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.Text = strData
    DecodeBase64 = Stream_BinaryToString(objNode.nodeTypedValue)
  
    ' Clean up
    Set objNode = Nothing
    Set objXML = Nothing
End Function

' Code by Tim Hastings
Function Stream_BinaryToString(Binary)
     Const adTypeText = 2
     Const adTypeBinary = 1
     
     'Create Stream object
     Dim BinaryStream 'As New Stream
     Set BinaryStream = CreateObject("ADODB.Stream")
     
     'Specify stream type - we want To save binary data.
     BinaryStream.Type = adTypeBinary
     
     'Open the stream And write binary data To the object
     BinaryStream.Open
     BinaryStream.Write Binary
     
     'Change stream type To text/string
     BinaryStream.Position = 0
     BinaryStream.Type = adTypeText
     
     'Specify charset For the output text (unicode) data.
     BinaryStream.Charset = "us-ascii"
     
     'Open the stream And get text/string data from the object
     Stream_BinaryToString = BinaryStream.ReadText
     Set BinaryStream = Nothing
End Function

Sub ReportAnySolutionSubOptimality()
          ' Call this after SolveModel to report to the user what happened. A dialog is shown if the optimal result was not found
15910     If SolveStatus <> OpenSolverResult.Optimal Then
15920         If SolveStatus = OpenSolverResult.AbortedThruUserAction Then
                  ' Do nothing; user has already interacted with some dialogs
15930         ElseIf LinearSolveStatusString = "" Then
                  ' We found an error without running CBC. LinearSolutionWasLoaded will be false. We expect to have a SolveStatusComment
15940             MsgBox "OpenSolver could not find an optimal solution, and reported:" _
                        & vbCrLf & SolveStatusString & vbCrLf & vbCrLf _
                       & IIf(SolveStatusComment = "", "", SolveStatusComment & vbCrLf & vbCrLf) _
                       & "No solution was available to load into the spreadsheet." _
                       , "OpenSolver: " & SolveStatusString
15950         Else
15960             MsgBox "OpenSolver could not find an optimal solution, and reported:" & vbCrLf _
                        & SolveStatusString & vbCrLf & vbCrLf _
                       & IIf(SolveStatusComment = "", "", SolveStatusComment & vbCrLf & vbCrLf) _
                       & IIf(LinearSolutionWasLoaded, "The solution generated has been loaded into the spreadsheet.", "No solution was available to load into the spreadsheet.") & vbCrLf & vbCrLf _
                       & " Solver reported: " & LinearSolveStatusString & "" _
                      , , "OpenSolver: " & SolveStatusString
15970         End If
15980     End If
End Sub

Function StrEx(d As Double) As String
      ' Convert a double to a string, always with a + or -. Also ensure we have "0.", not just "." for values between -1 and 1
              Dim s As String, prependedZero As String, sign As String
15990         s = Mid(str(d), 2)  ' remove the initial space (reserved by VB for the sign)
16000         prependedZero = IIf(left(s, 1) = ".", "0", "")  ' ensure we have "0.", not just "."
16010         sign = IIf(d >= 0, "+", "-")
16020         StrEx = sign + prependedZero + s
End Function

Sub WriteConstraintListToSheet(r As Range)
          ' Write a list of all the constraints in a column at cell r
          ' TODO: This will not correctly handle constraints on another sheet
          
16030     r.Cells(1, 1).Value2 = "Cons"
16040     r.Cells(1, 2).Value2 = "SP"
16050     r.Cells(1, 3).Value2 = "Inc"
16060     r.Cells(1, 4).Value2 = "Dec"
          
          Dim constraint As Long, row As Long, instance As Long
16070     constraint = 1
16080     For row = 1 To NumRows
16090         GetConstraintFromRow row, constraint, instance  ' Which Excel constraint are we in, and which instance?
              
              Dim UnusedConstraint As Boolean
16100         UnusedConstraint = SparseA(row).Count = 0
              
              Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
16110         GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
16120         If Not RHSCellRange Is Nothing Then
16130             RHSstring = RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
16140         Else
16150             RHSstring = RemoveActiveSheetNameFromString(RHSstring)  ' Make this formula or value prettier
16160             RHSstring = Replace(RHSstring, "$", "")
16170         End If
              
              Dim s As String
16180         s = IIf(UnusedConstraint, "", "") & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & _
                  SolverRelationAsString(Relation(row)) & RHSstring & IIf(UnusedConstraint, "", "")
          
16190         r.Cells(row + 1, 1).value = s
16200         r.Cells(row + 1, 2).Value2 = ShadowPrice(row)
16210         r.Cells(row + 1, 3).Value2 = IncreaseCon(row)
16220         r.Cells(row + 1, 4).Value2 = DecreaseCon(row)
16230     Next row

          'Write the variable duals
16240     row = row + 2
16250     r.Cells(row, 1).Value2 = "Vars"
16260     r.Cells(row, 2).Value2 = "RC"
16270     r.Cells(row, 3).Value2 = "Inc"
16280     r.Cells(row, 4).Value2 = "Dec"
16290     row = row + 1
          Dim i As Integer
16300     For i = 1 To UBound(ReducedCosts)
16310         r.Cells(row, 1).Value2 = VarCell(i)
16320         r.Cells(row, 2).Value2 = ReducedCosts(i)
16330         r.Cells(row, 3).Value2 = IncreaseVar(i)
16340         r.Cells(row, 4).Value2 = DecreaseVar(i)
16350         row = row + 1
16360     Next i

End Sub

Sub WriteModelToSheet()
          Dim oldCalculationMode As Integer
16370     oldCalculationMode = Application.Calculation
16380     Application.Calculation = xlCalculationManual
          
          'TODO : Put in error handling etc
16390     Application.StatusBar = "OpenSolver: Writing Model to worksheet 'Model'... " & numVars & " vars, " & NumRows & " rows."

          Dim var As Long, i As Long, row As Long, coeff As Double
          
16400     For var = 1 To numVars
16410         Worksheets("Model").Cells(1, 1 + var) = VarNames(var)
16420     Next var
          
          Dim c As Range
16430     var = 0
          
16440     For Each c In AdjustableCells
16450             var = var + 1
16460             If Not Intersect(c, IntegerCellsRange) Is Nothing Then
16470                 Worksheets("Model").Cells(2, 1 + var) = "Int"
16480             ElseIf Not Intersect(c, BinaryCellsRange) Is Nothing Then
16490                 Worksheets("Model").Cells(2, 1 + var) = "Bin"
16500             End If
16510     Next c
           
      '     For var = 1 To NumVars
      '        If VarTypes(i) = RelationInt Then
      '            Worksheets("Model").Cells(2, 1 + var) = "Int"
      '        ElseIf VarTypes(i) = RelationBin Then
      '            Worksheets("Model").Cells(2, 1 + var) = "Bin"
      '        End If
      '    Next var
      '

16520     For var = 1 To numVars
16530         Worksheets("Model").Cells(3, 1 + var) = CostCoeffs(var)
16540     Next var
              'For var = 1 To NumVars
              '    If a(row, var) <> 0 Then
              '        Worksheets("Model").Cells(4 + row, 1 + var) = a(row, var)
              '    End If
              'Next var
16550     With SparseA(row)
16560         For i = 1 To .Count
16570             var = .Index(i)
16580             coeff = .Coefficient(i)
16590             Worksheets("Model").Cells(4 + row, 1 + var) = coeff
16600         Next i
16610     End With
16620     For row = 1 To NumRows
16630         Worksheets("Model").Cells(4 + row, 1 + numVars + 1) = SolverRelationAsUnicodeChar(Relation(row))
16640         Worksheets("Model").Cells(4 + row, 1 + numVars + 2) = RHS(row)
16650     Next row
          
ExitSub:
16660     Application.StatusBar = False ' Resume normal status bar behaviour
16670     Application.ScreenUpdating = True
16680     Application.Calculation = oldCalculationMode
16690     Exit Sub
          
errorHandler:
16700     If Err.Number = 18 Then
16710         If MsgBox("You have pressed the Escape key. Do you wish to cancel?", _
                         vbCritical + vbYesNo + vbDefaultButton1, _
                         "OpenSolver: User Interrupt Occured...") = vbNo Then
16720             Resume 'continue on from where error occured
16730         Else
16740             Resume ExitSub
16750         End If
16760     End If
16770     MsgBox "OpenSolver encountered error " & Err.Number & ":" & vbCrLf & Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")") & vbCrLf & "Source = " & Err.Source, , "OpenSolver Code Error"
16780     Resume ExitSub
End Sub

' See http://www.decisionmodels.com/calcsecretsh.htm
'Application.CheckAbort ([KeepAbort])

' This clears any quick-solve initialisation; call this whenever the parameter range changes.
Sub ClearQuickSolve()
16790     Set QuickSolve = Nothing
End Sub

Function InitializeQuickSolve() As Boolean
16800     InitializeQuickSolve = False
16810     If ModelStatus <> ModelStatus_Built Then
16820         MsgBox "QuickSolve Initialisation Failed: The model cannot be solved as it has not yet been built.", , "OpenSolver" & sOpenSolverVersion & " Error"
16830         Exit Function
16840     End If

16850     Set QuickSolve = New CQuickSolve
16860     If QuickSolve.AnalyseParameters(Me, sheetName, NumConstraints, RHS, LHSRange, RHSRange, RHSFormula, RHSType, LHSOriginalValues, RHSOriginalValues) Then
16870         InitializeQuickSolve = True
16880     Else
16890         Set QuickSolve = Nothing
16900     End If
End Function

Function CanDoQuickSolveForActiveSheet() As Boolean
          ' Confirms that we can do a quicksolve for the active sheet because that is the sheet we are set up for
16910     CanDoQuickSolveForActiveSheet = False
16920     If QuickSolve Is Nothing Then
16930         MsgBox "Error: The quick solve data has not been initialized. Please choose Intialize Quick Solve, and then try again.", , "OpenSolver" & sOpenSolverVersion & " Error"
16940     ElseIf (ActiveWorkbook.Name <> RawWorkBookName) Or (ActiveSheet.Name <> RawSheetName) Then
16950         MsgBox "Error: The quick solve data has not been initialized for this worksheet. Please choose Intialize Quick Solve, and then try again.", , "OpenSolver" & sOpenSolverVersion & " Error"
16960     Else
16970         CanDoQuickSolveForActiveSheet = True
16980     End If
End Function

Function DoQuickSolve() As Boolean
          ' Update a previously built model
16990     DoQuickSolve = False
17000     If QuickSolve Is Nothing Then
17010         MsgBox "Error: The quick solve data has not been initialized. Please choose Intialize Quick Solve, and then try again.", , "OpenSolver" & sOpenSolverVersion & " Error"
17020     Else
17030         If QuickSolve.UpdateModelRHS(sheetName, RHS) Then
17040             If SolveModel(False) Then
17050                 DoQuickSolve = True
17060             End If
17070         End If
17080     End If
End Function

'Iain dunning
Public Function TestKeyExists(ByRef col As Collection, key As String) As Boolean
          
          'MsgBox Key
    On Error GoTo doesntExist:
          Dim Item As Variant
          
17090     Set Item = col(key)
          
17100     TestKeyExists = True
17110     Exit Function
          
doesntExist:
17120     If Err.Number = 5 Then
17130         TestKeyExists = False
17140     Else
17150         TestKeyExists = True
17160     End If
          
End Function

Sub HighlightNonLinearities(RowIsNonLinear() As Boolean, ObjectiveIsNonLinear As Boolean)
          ' Highlight all constraints (and the objective) that are non-linear using our standard model highlighting, but showing only individual cells, not ranges
          Dim constraint As Long, row As Long, instance As Long
17170      If SheetHasOpenSolverHighlighting(ActiveSheet) Then
17180         HideSolverModel
17190     End If
17200     DeleteOpenSolverShapes ActiveSheet
17210     InitialiseHighlighting
17220     constraint = 1
17230     For row = 1 To NumRows
17240         If RowIsNonLinear(row) Then
                  Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
17250             GetConstraintFromRow row, constraint, instance
17260             GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
                  'If Range(SheetName & "solver_lhs" & Constraint).Worksheet.Name <> ActiveWorkbook.ActiveSheet.Name Then
                  '    Set currentsheet = Range(SheetName & "solver_lhs" & Constraint).Worksheet
                  'Else
                  '    Set currentsheet = ActiveSheet
                  'End If
17270             RHSstring = StripWorksheetNameAndDollars(RHSstring, LHSCellRange.Worksheet) ' Strip any worksheet name and $'s from the RHS (useful if it is a formula)
                  Dim RHSisRange As Boolean
17280             RHSisRange = RHSType(constraint) = MultiCellRange Or RHSType(constraint) = SingleCellRange
17290             HighlightConstraint LHSCellRange.Worksheet, LHSCellRange, RHSisRange, RHSCellRange, RHSstring, Relation(row), 0  ' Show either a value or a formula
17300         End If
17310     Next row
17320     If ObjectiveIsNonLinear Then
              Dim ObjType As ObjectiveSenseType, temp As Integer, ObjectiveTargetValue As Double
17330         ObjType = UnknownObjectiveSense
17340         If GetNamedIntegerIfExists(book, sheetName & "solver_typ", temp) Then ObjType = temp
17350         If ObjType = TargetObjective Then GetNamedNumericValueIfExists book, sheetName & "solver_val", ObjectiveTargetValue   ' Target will be 0 if this fails
17360         AddObjectiveHighlighting ObjRange, ObjType, ObjectiveTargetValue
17370     End If
End Sub

Sub CheckLinearityOfModel()
          'this sub runs a linearity check on the model after its been solved.
          'Kathleen Gilbert - added version 1.1
          Dim row As Long, i As Long, RowIsNonLinear() As Boolean
          Dim ValueZero() As CIndexedCoeffs, ValueOne() As CIndexedCoeffs, ValueTen() As CIndexedCoeffs, OriginalSolutionValues() As Variant
          Dim nonlinearinformation As String
          Dim ObjectiveCoeffsZero() As Double, ObjectiveCoeffsOne() As Double, ObjectiveCoeffsTen() As Double
          Dim ObjectiveFunctionConstantZero As Double, ObjectiveFunctionConstantOne As Double, ObjectiveFunctionConstantTen As Double
17380     ReDim SolutionValues(numVars) As Double
17390     If NumRows > 0 Then ReDim Preserve ValueZero(NumRows) As CIndexedCoeffs
          
17400     nonlinearinformation = ""
          
          ' Remember the original decision variable values (in a variant array to handle multiple areas)
17410     OriginalSolutionValues = GetRangeValues(AdjustableCells)
          
17420     ReDim ObjectiveCoeffsZero(numVars) As Double, ObjectiveCoeffsOne(numVars) As Double, ObjectiveCoeffsTen(numVars) As Double
17430     If NumRows > 0 Then ReDim RowIsNonLinear(NumRows) As Boolean
          'Build each matrix where the decision variables start at zero (ValueZero()), one (ValueOne()) and ten (ValueTen())
17440     For row = 1 To NumRows
17450         Set ValueZero(row) = SparseA(row).Clone
17460     Next row
17470     For i = 1 To numVars
17480         ObjectiveCoeffsZero(i) = CostCoeffs(i)
17490     Next i
17500     ObjectiveFunctionConstantZero = ObjectiveFunctionConstant
          
17510     BuildModelFromSolverData 1
17520     If NumRows > 0 Then ReDim Preserve ValueOne(NumRows) As CIndexedCoeffs
17530     For row = 1 To NumRows
17540         Set ValueOne(row) = SparseA(row).Clone
17550     Next row
17560     For i = 1 To numVars
17570         ObjectiveCoeffsOne(i) = CostCoeffs(i)
17580     Next i
17590     ObjectiveFunctionConstantOne = ObjectiveFunctionConstant
          
17600     BuildModelFromSolverData 10
17610     If NumRows > 0 Then ReDim Preserve ValueTen(NumRows) As CIndexedCoeffs
17620     For row = 1 To NumRows
17630         Set ValueTen(row) = SparseA(row).Clone
17640     Next row
17650     For i = 1 To numVars
17660         ObjectiveCoeffsTen(i) = CostCoeffs(i)
17670     Next i
17680     ObjectiveFunctionConstantTen = ObjectiveFunctionConstant
          
          Dim constraint As Long, instance As Long
17690     constraint = 1
          
          'TODO: These tests should not have an AND, and is the model build code valid if we just shift the zero point?
          
          Dim NumEntries As Long, ValueZeroCounter As Long, ValueOneCounter As Long, ValueTenCounter As Long
          Dim FirstVar As Boolean, NonLinearityCount As Long
          'Go through each row and check each coefficient individually. if it is not within the tolerance the its nonlinear
17700     For row = 1 To NumRows
17710         RowIsNonLinear(row) = False
              
              'This is used to display the constriant
17720         FirstVar = True
17730         ValueZeroCounter = ValueZero(row).Count
17740         ValueOneCounter = ValueOne(row).Count
17750         ValueTenCounter = ValueTen(row).Count
              'find out how many variables it dependent on
17760         NumEntries = ChooseMaximum(ValueZeroCounter, ValueOneCounter, ValueTenCounter)
17770         For i = 1 To NumEntries
17780             If TestExistanceOfEntry(ValueZero(row), ValueOne(row), ValueTen(row), i) Then
                      'do a ratio test
17790                 If Abs(ValueOne(row).Coefficient(i) - ValueZero(row).Coefficient(i)) / (1 + Abs(ValueOne(row).Coefficient(i))) > EPSILON _
                      And Abs(ValueOne(row).Coefficient(i) - ValueTen(row).Coefficient(i)) / (1 + Abs(ValueOne(row).Coefficient(i))) > EPSILON Then
17800                     GetConstraintFromRow row, constraint, instance
17810                     If NonLinearityCount <= 10 Then AddNonLinearInfoToString ValueOne(row).Index(i), nonlinearinformation, FirstVar, constraint, instance
17820                     FirstVar = False
17830                     RowIsNonLinear(row) = True
17840                 End If
17850             Else
17860                 GetConstraintFromRow row, constraint, instance
                      Dim VariableIndex As Long
17870                 VariableIndex = GetEntry(ValueZero(row), ValueOne(row), ValueTen(row), i)
17880                 If NonLinearityCount <= 10 Then AddNonLinearInfoToString VariableIndex, nonlinearinformation, FirstVar, constraint, instance
17890                 FirstVar = False
17900                 RowIsNonLinear(row) = True
17910             End If
17920         Next i
17930         If RowIsNonLinear(row) Then NonLinearityCount = NonLinearityCount + 1
17940     Next row
          
17950     If NonLinearityCount > 10 Then
17960         nonlinearinformation = nonlinearinformation & vbNewLine & " and " & CStr(NonLinearityCount - 10) & " other instances."
17970     End If

          Dim ObjectiveIsNonLinear As Boolean
17980     ObjectiveIsNonLinear = False
17990     For i = 1 To numVars
18000         If Abs(ObjectiveCoeffsZero(i) - ObjectiveCoeffsOne(i)) / (1 + Abs(ObjectiveCoeffsZero(i))) > EPSILON _
              And Abs(ObjectiveCoeffsOne(i) - ObjectiveCoeffsTen(i)) / (1 + Abs(ObjectiveCoeffsOne(i))) > EPSILON Then
18010             If Not ObjectiveIsNonLinear Then
18020                 ObjectiveIsNonLinear = True
18030                 nonlinearinformation = nonlinearinformation & vbNewLine & vbNewLine & "The objective function is nonlinear in the following variables: " & VarNames(i)
18040             Else
18050                 nonlinearinformation = nonlinearinformation & " , " & VarNames(i)
18060             End If
                   
18070         End If
18080     Next i
          
          'Put the solution back on the sheet
18090     SetRangeValues AdjustableCells, OriginalSolutionValues
          
          'display dialog to user
18100     NonlinearForm.TextBox2.Caption = TrimBlankLines(nonlinearinformation)
18110     If nonlinearinformation = "" Then
18120         NonlinearForm.TextBox2.Caption = "There have been no instances of nonlinearity found in this model. Some models can generate warnings of non-linearity " _
                                             & "because of numerical errors that accumulate in the spreadsheet. OpenSolver's non-linearity check can be disabled under OpenSolver's " _
                                             & "Options settings."
18130     End If
18140     NonlinearForm.TextBox2.TextAlign = fmTextAlignLeft
18150     NonlinearForm.TextBox2.AutoSize = False
18160     NonlinearForm.TextBox2.height = 20
18170     NonlinearForm.TextBox2.AutoSize = True
18180     NonlinearForm.TextBox2.AutoSize = False
18190     If NonlinearForm.TextBox2.height > 250 Then NonlinearForm.TextBox2.height = 250
18200     NonlinearForm.TextBox2.width = 546
          ' NonlinearForm.TextBox2.top = 1
18210     NonlinearForm.width = NonlinearForm.TextBox2.width + 2 * NonlinearForm.TextBox2.left
          'NonlinearForm.TextBox2.left = NonlinearForm.TextBox2.width * 0.1
18220     NonlinearForm.FullCheck.Visible = False
18230     NonlinearForm.HighlightBox.Visible = True
          'NonlinearForm.ContinueButton.left = Int(NonlinearForm.TextBox2.left + NonlinearForm.TextBox2.width * 0.5 - (NonlinearForm.ContinueButton.width / 2))
          'NonlinearForm.HighlightBox.left = NonlinearForm.TextBox2.left
18240     NonlinearForm.HighlightBox.top = NonlinearForm.TextBox2.top + NonlinearForm.TextBox2.height + 5
          'NonlinearForm.FullCheck.left = NonlinearForm.TextBox2.left
          ' NonlinearForm.FullCheck.top = NonlinearForm.HighlightBox.top + NonlinearForm.HighlightBox.height + 5
18250     NonlinearForm.ContinueButton.top = NonlinearForm.HighlightBox.top  ' Enought space around check box
18260     NonlinearForm.height = NonlinearForm.ContinueButton.top + NonlinearForm.ContinueButton.height + 30
18270     NonlinearForm.Show
          
18280     If NonlinearForm.HighlightBox.value = True Then
18290         HighlightNonLinearities RowIsNonLinear, ObjectiveIsNonLinear
18300     End If
End Sub

Function TestExistanceOfEntry(ValueZero, ValueOne, ValueTen, i) As Boolean
          'Check if ALL the index exists
18310     On Error GoTo errorHandler
18320     If ValueZero.Index(i) <> 0 And ValueOne.Index(i) <> 0 And ValueTen.Index(i) <> 0 Then
18330         TestExistanceOfEntry = True
18340     End If
errorHandler:
18350     If Err.Number = 9 Then
18360         TestExistanceOfEntry = False
18370     End If
End Function

Function GetEntry(ValueZero, ValueOne, ValueTen, i) As Long
          'Return the i'th entry from any one of these arrays; it may not exist in all of them
18380     On Error Resume Next
18390     If i <= ValueZero.Count Then
18400         GetEntry = ValueZero.Index(i)
18410     ElseIf i <= ValueOne.Count Then
18420         GetEntry = ValueOne.Index(i)
18430     ElseIf i <= ValueTen.Count Then
18440         GetEntry = ValueTen.Index(i)
18450     End If
End Function

Sub AddNonLinearInfoToString(var As Long, nonlinearinformation As String, FirstVar As Boolean, constraint As Long, instance As Long)
          'This function is used by the linearity checker to write the string storing all the nonlinear information
18460     If FirstVar = True Then
18470         If LHSType(constraint) = SolverInputType.SingleCellRange Then
18480             nonlinearinformation = nonlinearinformation & vbNewLine & "In the constraint: " & ConstraintSummary(constraint) & "," & vbNewLine & "  the model appears to be non-linear in the decision variables: " & VarNames(var)
18490         Else
                  ' Dim LHSValue As Double, RHSValue As Double, failedcase As Long, sLHS As String, sRHS As String
                  ' GetExcelConstraintInfoFromRow Constraint, Row, LHSValue, RHSValue, failedcase, sLHS, sRHS
18500             nonlinearinformation = nonlinearinformation & vbNewLine & "In instance " & instance & " of the constraint: " & ConstraintSummary(constraint) & "," & vbNewLine & "the model appears to be non-linear in the following decision variables: " & VarNames(var)
18510         End If
18520     Else
18530         nonlinearinformation = nonlinearinformation & ", " & VarNames(var)
18540     End If
End Sub

Function ChooseMaximum(EntriesZero As Long, EntriesOne As Long, EntriesTen As Long) As Long
          'This function finds which of three numbers is the largest
18550     If EntriesZero >= EntriesOne Then
18560         If EntriesZero >= EntriesTen Then
18570             ChooseMaximum = EntriesZero
18580         ElseIf EntriesTen >= EntriesZero Then
18590             ChooseMaximum = EntriesTen
18600         End If
18610     ElseIf EntriesTen > EntriesOne Then
18620         ChooseMaximum = EntriesTen
18630     Else
18640         ChooseMaximum = EntriesOne
18650     End If
          
End Function

Sub QuickLinearityCheck(fullLinearityCheckWasPerformed As Boolean)
          'This checks the value we would expect against what was reported for linearity
          ' Returns false if a full check was performed by the user, meaning the model result is no longer valid.
          'Kathleen Gilbert - added version 1.1
          
18660     On Error GoTo 0
18670     fullLinearityCheckWasPerformed = False
          
          Dim row As Long ', i As Long
          Dim NonLinearInfo As String, NonlinearConstraint As Boolean
          'Dim x As Double,
          Dim ExpectedValue As Double, SolutionValue As Double, SolutionValueLHS As Double, SolutionValueRHS As Double
          Dim constraint As Long, i As Long, j As Long
          'Dim sLHS As String, sRHS As String
          'Dim LHSArray As Variant
          Dim CurrentLHSValues As Variant, CurrentRHSValues As Variant
          Dim RowIsNonLinear() As Boolean
18680     If NumRows > 0 Then ReDim RowIsNonLinear(NumRows) As Boolean
          
18690     If Not ForceCalculate("Warning: The worksheet calculation did not complete during the linearity test, and so the test may not be correct. Would you like to retry?") Then
18700         Exit Sub
18710     End If
          
          ' Get all the decision variable values off the sheet
          Dim DecisionVariableValues() As Double
18720     ReDim DecisionVariableValues(numVars)
18730     DecisionVariableValues = GetDecisionVariableValuesOffSheet
          
18740     NonLinearInfo = ""
          Dim NonLinearityCount As Long
          
18750     row = 1
18760     For constraint = 1 To NumConstraints
18770         If Not LHSRange(constraint) Is Nothing Then ' Skip INT and BINARY constraint
                  ' Get current value(s) for LHS and RHS of this constraint off the sheet
18780             GetCurrentConstraintValues constraint, CurrentLHSValues, CurrentRHSValues
                  
18790             If RHSType(constraint) <> SolverInputType.MultiCellRange Then
18800                 SolutionValueRHS = CurrentRHSValues
18810             End If
                  
                  Dim instance As Long
18820             instance = 0
18830             For i = 1 To UBound(CurrentLHSValues, 1)
18840                 For j = 1 To UBound(CurrentLHSValues, 2)
18850                     instance = instance + 1
18860                     SolutionValueLHS = CurrentLHSValues(i, j)
18870                     If RHSType(constraint) = SolverInputType.MultiCellRange Then
                              '---------------------------------------------------------------
                              'Check whether the LHS and RHS are parallel or perpendicular
                              '---------------------------------------------------------------
18880                         If UBound(CurrentLHSValues, 1) = UBound(CurrentRHSValues, 1) Then
18890                             SolutionValueRHS = CurrentRHSValues(i, j)
18900                         Else
18910                             SolutionValueRHS = CurrentRHSValues(j, i)
18920                         End If
18930                     End If
18940                     SolutionValue = SolutionValueLHS - SolutionValueRHS
                      
                          'Find out what we expect the value to be from Ax = b. We track the maximum value we encounter during the calculation
                          'so that we have some idea of the errors we might expect
                          Dim maxValueInCalculation As Double
18950                     maxValueInCalculation = 0
18960                     ExpectedValue = SparseA(row).Evaluate_RecordPrecision(DecisionVariableValues, maxValueInCalculation) - RHS(row)
18970                     If Abs(RHS(row)) > maxValueInCalculation Then maxValueInCalculation = Abs(RHS(row))
          
                          ' do a ratio test
18980                     If Abs(ExpectedValue - SolutionValue) / (1 + Abs(ExpectedValue)) > Max_Double(EPSILON, EPSILON * maxValueInCalculation) Then
18990                         If NonLinearInfo = "" Then NonLinearInfo = "The following constraint(s) do not appear to be linear: "
19000                         If NonLinearityCount <= 10 Then
19010                             NonLinearInfo = NonLinearInfo & vbNewLine & ConstraintSummary(constraint)
                              
                                  Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
19020                             GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
                                  ' If the RHS is a range, we show its address; if not, RHSString contains the RHS's constant or formula
19030                             If Not RHSCellRange Is Nothing Then RHSstring = RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
19040                             NonLinearInfo = NonLinearInfo & ": instance " _
                                         & instance _
                                         & ", LHS=" & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) _
                                         & ", RHS=" & RHSstring _
                                         & ", " & ExpectedValue _
                                         & "<>" & SolutionValue
                              
19050                         End If
19060                         NonLinearityCount = NonLinearityCount + 1
19070                         RowIsNonLinear(row) = True
19080                         NonlinearConstraint = True
19090                     End If
                    
19100                     row = row + 1
19110                 Next j
19120             Next i
19130         End If
19140     Next constraint
19150     If NonLinearityCount > 10 Then
19160         NonLinearInfo = NonLinearInfo & vbNewLine & " and " & CStr(NonLinearityCount - 10) & " other constraints."
19170     End If

          'check objective function for linearity
          Dim CalculatedObjValue As Double, ObservedObjValue As Double, ObjectiveIsNonLinear As Boolean
19180     On Error Resume Next
          Dim NM As Name
19190     Set NM = book.Names(sheetName & "solver_opt")
19200     If Not Err.Number <> 0 Then
19210         ObservedObjValue = ObjRange.Value2
19220    End If
19230     On Error GoTo 0
19240     CalculatedObjValue = CalcObjFnValue(DecisionVariableValues)
19250     ObjectiveIsNonLinear = Abs(CalculatedObjValue - ObservedObjValue) / (1 + Abs(CalculatedObjValue)) > EPSILON
19260     If ObjectiveIsNonLinear Then
19270        NonLinearInfo = "The objective function is not linear." & vbNewLine & vbNewLine & NonLinearInfo
19280     End If
          
          
          'Set the userform up and display any information on nonlinear constraints
19290     If NonLinearInfo <> "" Then
19300         NonlinearForm.TextBox2.Caption = "WARNING : " & vbNewLine & TrimBlankLines(NonLinearInfo)
              
              'If its only the objective function hide the full check
              'If NonlinearConstraint = True Then
19310            NonlinearForm.FullCheck.Visible = True
              'Else
              '    NonlinearForm.FullCheck.Visible = False
              '    NonlinearForm.FullCheck.value = False
              'End If
              
              'formatting of the user form NonlinearForm.TextBox2.AutoSize = True
19320         NonlinearForm.TextBox2.AutoSize = False
19330         NonlinearForm.TextBox2.height = 20
19340         NonlinearForm.TextBox2.AutoSize = True
19350         NonlinearForm.TextBox2.AutoSize = False
19360         If NonlinearForm.TextBox2.height > 250 Then NonlinearForm.TextBox2.height = 250
19370         NonlinearForm.TextBox2.width = 546
19380         NonlinearForm.width = NonlinearForm.TextBox2.width + 2 * NonlinearForm.TextBox2.left
              'NonlinearForm.TextBox2.left = NonlinearForm.TextBox2.width * 0.1
19390         NonlinearForm.TextBox2.TextAlign = fmTextAlignLeft
19400         NonlinearForm.HighlightBox.Visible = True
19410         NonlinearForm.FullCheck.Caption = "Run a full linearity check. (This will destroy the current solution) "
19420         NonlinearForm.Caption = "OpenSolver: Linearity check "
19430         NonlinearForm.HighlightBox.Caption = "Highlight the nonlinearities"
              'NonlinearForm.HighlightBox.left = NonlinearForm.TextBox2.left
              'NonlinearForm.FullCheck.left = NonlinearForm.TextBox2.left
19440         NonlinearForm.HighlightBox.top = NonlinearForm.TextBox2.height + NonlinearForm.TextBox2.top + 5
19450         NonlinearForm.FullCheck.top = NonlinearForm.HighlightBox.top + NonlinearForm.HighlightBox.height + 0 ' Enought space around check box anyway
              ' NonlinearForm.ContinueButton.top = NonlinearForm.FullCheck.top + NonlinearForm.FullCheck.height + 3 ' Enough space around check box anyway
19460         NonlinearForm.ContinueButton.top = NonlinearForm.HighlightBox.top + 6 ' Enough space around check box anyway
              'NonlinearForm.ContinueButton.left = NonlinearForm.TextBox2.width * 0.4
19470         NonlinearForm.height = NonlinearForm.ContinueButton.top + NonlinearForm.ContinueButton.height + 30
              'NonlinearForm.height = (NonlinearForm.TextBox2.top + NonlinearForm.TextBox2.height + NonlinearForm.FullCheck.height _
              '        + NonlinearForm.HighlightBox.height + NonlinearForm.ContinueButton.height + NonlinearForm.ContinueButton.top) * 0.8
19480         NonlinearForm.Show
              
              'showing the nonlinear constraints
19490         If NonlinearForm.HighlightBox.value = True Then
19500             HighlightNonLinearities RowIsNonLinear, ObjectiveIsNonLinear
19510         End If
19520         If NonlinearForm.FullCheck.value = True Then
                  'Full linearity check run
19530             CheckLinearityOfModel
19540             fullLinearityCheckWasPerformed = True
19550         End If
19560     End If
          
End Sub
' For code for parsing Excel formulae, see
' http://ewbi.blogs.com/develops/2007/03/excel_formula_p.html
' http://ewbi.blogs.com/develops/2004/12/excel_formula_p.html
' http://www.dailydoseofexcel.com/archives/2005/04/01/excel-formula-formatter/
' http://homepages.ecs.vuw.ac.nz/~elvis/db/FishBrainWiki?Excel
' http://www.nsl.com/papers/excel.htm
' http://www.dicks-blog.com/archives/2005/04/01/excel-formula-formatter/
' http://vangelder.orconhosting.net.nz/excel/audxl.html
' For a .net Excel-like formulae parser, see
' http://www.codeproject.com/KB/vb/FormulaEngine.aspx
' http://www.codeproject.com/KB/office/Excel_DataReader.aspx

'Function CorrectWorksheetCalcOccured() As Boolean
'    If Application.CalculationState <> xlDone Then Application.Calculate
'        If Application.CalculationState <> xlDone Then Application.Calculate
'        If Application.CalculationState <> xlDone Then
'            DoEvents
'            Application.CalculateFullRebuild
'            DoEvents
'        End If
'        Do While Application.CalculationState <> xlDone
'            If MsgBox("Warning: The worksheet calculation did not complete during the linearity test, and so the model may not be generated correctly. Would you like to retry?", _
'                        vbCritical + vbRetryCancel + vbDefaultButton1, _
'                        "OpenSolver: Calculation Error Occured...") = vbCancel Then
'
'               Exit Do
'            Else 'Recalculate the workbook if the user wants to retry
'                Application.Calculate
'            End If
'        Loop
'        If Application.CalculationState <> xlDone Then
'            CorrectWorksheetCalcOccured = False
'        Else
'            CorrectWorksheetCalcOccured = True
'        End If
'End Function
Private Sub Class_Initialize()

End Sub

Function ReadGurobiModel(SolutionFilePathName As String, bGetDuals As Boolean, _
        ErrorString As String) As Boolean
          
19570     ReadGurobiModel = False
          Dim Line As String, Index As Integer
19580     On Error GoTo readError
          Dim solutionExpected As Boolean
19590     solutionExpected = True
          
19600     Open SolutionFilePathName For Input As 1 ' supply path with filename
19610     Line Input #1, Line
          'Get the returned status code from gurobi.
          'List of return codes can be seen at - http://www.gurobi.com/documentation/5.1/reference-manual/node865#sec:StatusCodes
19620     If Line = GurobiResult.Optimal Then
19630         SolveStatus = OpenSolverResult.Optimal
19640         SolveStatusString = "Optimal"
19650         LinearSolveStatus = LinearSolveResult.Optimal
19660     ElseIf Line = GurobiResult.Infeasible Then
19670         SolveStatus = OpenSolverResult.Infeasible
19680         SolveStatusString = "No Feasible Solution"
19690         solutionExpected = False
19700         LinearSolveStatus = LinearSolveResult.Infeasible
19710     ElseIf Line = GurobiResult.InfOrUnbound Then
19720         SolveStatus = OpenSolverResult.Unbounded
19730         SolveStatusString = "No Solution Found (Infeasible or Unbounded)"
19740         solutionExpected = False
19750         LinearSolveStatus = LinearSolveResult.Unbounded
19760     ElseIf Line = GurobiResult.Unbounded Then
19770         SolveStatus = OpenSolverResult.Unbounded
19780         SolveStatusString = "No Solution Found (Unbounded)"
19790         solutionExpected = False
19800         LinearSolveStatus = LinearSolveResult.Unbounded
19810     ElseIf Line = GurobiResult.SolveStoppedTime Then
19820         SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
19830         SolveStatusString = "Stopped on Time Limit"
19840         LinearSolveStatus = LinearSolveResult.SolveStopped
19850     ElseIf Line = GurobiResult.SolveStoppedIter Then
19860         SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
19870         SolveStatusString = "Stopped on Iteration Limit"
19880         LinearSolveStatus = LinearSolveResult.SolveStopped
19890     ElseIf Line = GurobiResult.SolveStoppedUser Then
19900         SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
19910         SolveStatusString = "Stopped on Ctrl-C"
19920         LinearSolveStatus = LinearSolveResult.SolveStopped
19930     ElseIf Line = GurobiResult.Unsolved Then
19940         SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
19950         SolveStatusString = "Stopped on Gurobi Numerical difficulties"
19960         LinearSolveStatus = LinearSolveResult.SolveStopped
19970     ElseIf Line = GurobiResult.SubOptimal Then
19980         SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
19990         SolveStatusString = "Unable to satisfy optimality tolerances; a sub-optimal solution is available."
20000         LinearSolveStatus = LinearSolveResult.SolveStopped
20010     Else
20020         ErrorString = "The response from the Gurobi solver is not recognised. The response was: " & Line
20030         GoTo readError
20040     End If
          
20050     If solutionExpected Then
20060         Line Input #1, Line  ' Optimal - objective value              22
20070         If Line <> "" Then
20080             Index = InStr(Line, "=")
                  Dim ObjectiveValue As Double
20090             ObjectiveValue = Val(Mid(Line, Index + 2))
                  Dim i As Integer
20100             i = 1
20110             While Not EOF(1)
20120                 Line Input #1, Line
20130                 Index = InStr(Line, " ")
20140                 ReDim Preserve FinalVarValue(i)
20150                 ReDim Preserve VarCell(i)
20160                 FinalVarValue(i) = Val(right(Line, Len(Line) - Index))
                      'Get the variable name
20170                 VarCell(i) = left(Line, Index - 1)
20180                 If left(VarCell(i), 1) = "_" Then VarCell(i) = Mid(VarCell(i), 2) ' Strip any _ character added to make a valid name
20190                 i = i + 1
20200             Wend
20210         End If
20220         AdjustableCells.Value2 = 0
              Dim NumVar As Integer, j As Integer
20230         NumVar = UBound(FinalVarValue, 1)
20240         For i = 1 To NumVar
20250             AdjustableCells.Worksheet.Range(VarCell(i)).Value2 = FinalVarValue(i)
20260         Next i
              
20270         If bGetDuals Then
20280             ReDim FinalValue(NumRows)
20290             ReDim ShadowPrice(NumRows)
20300             ReDim IncreaseCon(NumRows)
20310             ReDim DecreaseCon(NumRows)
20320             ReDim ReducedCosts(NumVar)
20330             ReDim IncreaseVar(NumVar)
20340             ReDim DecreaseVar(NumVar)
20350             Open Replace(SolutionFilePathName, "modelsolution", "sensitivityData") For Input As 2
                  Dim index2 As Integer
                  Dim Stuff() As String
20360             ReDim Stuff(3)
20370             For i = 1 To NumVar
20380                 Line Input #2, Line
20390                 For j = 1 To 3
20400                     index2 = InStr(Line, ",")
20410                     If index2 <> 0 Then
20420                         Stuff(j) = left(Line, index2 - 1)
20430                     Else
20440                         Stuff(j) = Line
20450                     End If
20460                     Line = Mid(Line, index2 + 1)
20470                 Next j
20480                 ReducedCosts(i) = Stuff(1)
20490                 IncreaseVar(i) = Stuff(3) - CostCoeffs(i)
20500                 DecreaseVar(i) = CostCoeffs(i) - Stuff(2)
20510             Next i
20520             ReDim Stuff(5)
20530             For i = 1 To NumRows
20540                 Line Input #2, Line
20550                 For j = 1 To 5
20560                     index2 = InStr(Line, ",")
20570                     If index2 <> 0 Then
20580                         Stuff(j) = left(Line, index2 - 1)
20590                     Else
20600                         Stuff(j) = Line
20610                     End If
20620                     Line = Mid(Line, index2 + 1)
20630                 Next j
20640                 ShadowPrice(i) = Stuff(1)
20650                 IncreaseCon(i) = Stuff(5) - Stuff(2)
20660                 DecreaseCon(i) = Stuff(2) - Stuff(4)
20670                 FinalValue(i) = Stuff(2) - Stuff(3)
20680             Next i
20690         End If
20700         ReadGurobiModel = True
20710     End If

exitFunction:
20720     Close #1
20730     Close #2
20740     Exit Function
          
readError:
20750     Close #1
20760     Close #2
End Function

Function ReadCBCModel(SolutionFilePathName As String, CommandLineRunString As String, bGetDuals As Boolean, rConstraintList As Range, _
                            ErrorString As String) As Boolean
                            
20770     ReadCBCModel = False
20780     Open SolutionFilePathName For Input As 1 ' supply path with filename
20790     Line Input #1, LinearSolveStatusString  ' Optimal - objective value              22
          ' Line Input #1, junk ' get rest of line
          Dim solutionExpected As Boolean
20800     solutionExpected = True
20810     If LinearSolveStatusString Like "Optimal*" Then
20820         SolveStatus = OpenSolverResult.Optimal
20830         SolveStatusString = "Optimal"
20840         LinearSolveStatus = LinearSolveResult.Optimal
              '
20850     ElseIf LinearSolveStatusString Like "Infeasible*" Then
20860         SolveStatus = OpenSolverResult.Infeasible
20870         SolveStatusString = "No Feasible Solution"
20880         LinearSolveStatus = LinearSolveResult.Infeasible
              '
20890     ElseIf LinearSolveStatusString Like "Integer infeasible*" Then
20900         SolveStatus = OpenSolverResult.Infeasible
20910         SolveStatusString = "No Feasible Integer Solution"
20920         LinearSolveStatus = LinearSolveResult.IntegerInfeasible
              '
20930     ElseIf LinearSolveStatusString Like "Unbounded*" Then
20940         SolveStatus = OpenSolverResult.Unbounded
20950         SolveStatusString = "No Solution Found (Unbounded)"
20960         LinearSolveStatus = LinearSolveResult.Unbounded
20970         solutionExpected = False
              '
20980     ElseIf LinearSolveStatusString Like "Stopped on time *" Then ' Stopped on iterations or time
20990         SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
21000         SolveStatusString = "Stopped on Time Limit"
21010         If LinearSolveStatusString Like "*(no integer solution - continuous used)*" Then
21020             SolveStatusString = SolveStatusString & ": No integer solution found. Fractional solution returned."
21030         End If
21040         LinearSolveStatus = LinearSolveResult.SolveStopped
              '
21050     ElseIf LinearSolveStatusString Like "Stopped on iterations*" Then ' Stopped on iterations or time
21060         SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
21070         SolveStatusString = "Stopped on Iteration Limit"
21080         If LinearSolveStatusString Like "*(no integer solution - continuous used)*" Then
21090             SolveStatusString = SolveStatusString & ": No integer solution found. Fractional solution returned."
21100         End If
21110         LinearSolveStatus = LinearSolveResult.SolveStopped
              '
21120     ElseIf LinearSolveStatusString Like "Stopped on difficulties*" Then ' Stopped on iterations or time
21130         SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
21140         SolveStatusString = "Stopped on CBC difficulties"
21150         If LinearSolveStatusString Like "*(no integer solution - continuous used)*" Then
21160             SolveStatusString = SolveStatusString & ": No integer solution found. Fractional solution returned."
21170         End If
21180         LinearSolveStatus = LinearSolveResult.SolveStopped
              '
21190     ElseIf LinearSolveStatusString Like "Stopped on ctrl-c*" Then ' Stopped on iterations or time
21200         SolveStatus = OpenSolverResult.TimeLimitedSubOptimal
21210         SolveStatusString = "Stopped on Ctrl-C"
21220         If LinearSolveStatusString Like "*(no integer solution - continuous used)*" Then
21230             SolveStatusString = SolveStatusString & ": No integer solution found. Fractional solution returned."
21240         End If
21250         LinearSolveStatus = LinearSolveResult.SolveStopped
              '
21260     ElseIf LinearSolveStatusString Like "Status unknown*" Then
21270         ErrorString = "CBC solver did not solve the problem, suggesting there was an error in the CBC input parameters. The response was: " & vbCrLf _
               & LinearSolveStatusString _
               & vbCrLf & "The CBC command line was:" _
               & vbCrLf & CommandLineRunString
21280         GoTo ExitSub
21290     Else
21300         ErrorString = "The response from the CBC solver is not recognised. The response was: " & LinearSolveStatusString
21310         GoTo ExitSub
21320     End If
          
          ' Remove the double spaces from LinearSolveStatusString
21330     LinearSolveStatusString = Replace(LinearSolveStatusString, "    ", " ")
21340     LinearSolveStatusString = Replace(LinearSolveStatusString, "   ", " ")
21350     LinearSolveStatusString = Replace(LinearSolveStatusString, "  ", " ")
21360     LinearSolveStatusString = Replace(LinearSolveStatusString, "  ", " ")
21370     LinearSolveStatusString = Replace(LinearSolveStatusString, "  ", " ")
21380     LinearSolveStatusString = Replace(LinearSolveStatusString, "  ", " ")

21390     If solutionExpected Then
              ' We read in whatever solution CBC returned
21400         Application.StatusBar = "OpenSolver: Loading Solution... " & LinearSolveStatusString
              ' Zero the current decision variables
21410         AdjustableCells.Value2 = 0
              ' Faster code; put a zero into first adjustable cell, and copy it to all the adjustable cells
              ' AdjustableCells.Cells(0, 0).Value = 0
              ' AdjustableCells.Cells(0, 0).Copy
              ' AdjustableCells.PasteSpecial xlPasteValues
          
              ' Read in the Solution File
              ' This gives the non-zero? variable values
              ' Lines like:       0 AZ70                  15                      0
              ' ...being? : Index Name Value ReducedCost
              Dim Line As String, SplitLine() As String, Index As Double, NameValue As String, value As Double, CBCConstraintIndex As Long
21420         If bGetDuals Then
                  Dim j As Integer, row As Integer, i As Integer
                  'Dim FinalValue() As String, ShadowPrice() As String
21430             ReDim FinalValue(NumRows)
21440             ReDim ShadowPrice(NumRows)
21450             j = 1
21460             CBCConstraintIndex = 0
21470             For row = 1 To NumRows
21480                 If SparseA(row).Count = 0 Then
                          ' This constraint was not written to the model, as it had no coefficients. Just ignore it.
21490                     rConstraintList.Cells(row, 2).ClearContents
21500                 Else
21510                     Line Input #1, Line
21520                     SplitLine = split(Line, " ")    ' 0 indexed; item 0 is the variable index
                          ' Skip over the blank items in the split (multiple delimiters give multiple items), getting the real items
21530                     i = 0
21540                     While SplitLine(i) = ""
21550                         i = i + 1
21560                     Wend
                          ' In the case of LpStatusInfeasible, we can get lines that start **. We strip the **
21570                     If SplitLine(i) = "**" Then i = i + 1
21580                     While SplitLine(i) = ""
21590                         i = i + 1
21600                     Wend
                          ' Get and check the index of the row
21610                     If Val(SplitLine(i)) <> CBCConstraintIndex Then
21620                         ErrorString = "While reading the CBC solution file, OpenSolver found an unexpected constraint row."
21630                         GoTo ExitSub
21640                     End If
21650                     i = i + 1
21660                     While SplitLine(i) = ""
21670                         i = i + 1
21680                     Wend
                          ' Get the constraint name; we don't use this
21690                     NameValue = SplitLine(i)
21700                     i = i + 1
21710                     While SplitLine(i) = ""
21720                         i = i + 1
21730                     Wend
21740                     FinalValue(j) = SplitLine(i)
                          ' Skip the constraint LHS value - we don't need this
21750                     i = i + 1
21760                     While SplitLine(i) = ""
21770                         i = i + 1
21780                     Wend
                          ' Get the dual value
21790                     If ObjectiveSense = MaximiseObjective Then
21800                         value = -1 * Val(SplitLine(i))
                              'rConstraintList.Cells(row, 2).Value2 = Value
21810                     Else
21820                         value = Val(SplitLine(i))
                              'rConstraintList.Cells(row, 2).Value2 = Value
21830                     End If
21840                     ShadowPrice(j) = value
21850                     If InStr(ShadowPrice(j), "E-16") Then
21860                         ShadowPrice(j) = "0"
21870                     End If
21880                     CBCConstraintIndex = CBCConstraintIndex + 1
21890                     j = j + 1
21900                 End If
21910             Next row
21920             Call ReadCBCSensitivityData(SolutionFilePathName)
21930         End If
            
              ' Now we read in the decision variable values
21940         j = 1
21950         While Not EOF(1)
21960             Line Input #1, Line
21970             SplitLine = split(Line, " ")    ' 0 indexed; item 0 is the variable index
                  ' Skip over the blank items in the split (multiple delimiters give multiple items), getting the real items
21980             i = 0
21990             While SplitLine(i) = ""
22000                 i = i + 1
22010             Wend
                  ' In the case of LpStatusInfeasible, we can get lines that start **. We strip the **
22020             If SplitLine(i) = "**" Then i = i + 1
22030             While SplitLine(i) = ""
22040                 i = i + 1
22050             Wend
                  ' Get the index of the variable
22060             Index = Val(SplitLine(i))
22070             i = i + 1
22080             While SplitLine(i) = ""
22090                 i = i + 1
22100             Wend
                  ' Get the variable name, stripping any leading "_"
22110             NameValue = SplitLine(i)
22120             If left(NameValue, 1) = "_" Then NameValue = Mid(NameValue, 2) ' Strip any _ character added to make a valid name
22130             i = i + 1
22140             While SplitLine(i) = ""
22150                 i = i + 1
22160             Wend
22170             ReDim Preserve FinalVarValue(j) As String
22180             FinalVarValue(j) = Val(SplitLine(i))
                  'Write to the sheet containing the decision variables (which may not be the active sheet)
22190             AdjustableCells.Worksheet.Range(NameValue).Value2 = FinalVarValue(j)
                 
                  'ConvertFullLPFileVarNameToRange(name, AdjCellsSheetIndex).Value2 = Value
22200             If bGetDuals Then
22210                 i = i + 1
22220                 While SplitLine(i) = ""
22230                     i = i + 1
22240                 Wend
22250                 If ObjectiveSense = MaximiseObjective Then
22260                     value = -1 * Val(SplitLine(i))
22270                 Else
22280                     value = Val(SplitLine(i))
22290                 End If
22300                 ReDim Preserve ReducedCosts(j) As String
22310                 ReDim Preserve VarCell(j) As String
22320                 ReducedCosts(j) = str(value)
22330                 If InStr(ReducedCosts(j), "E-16") Then
22340                     ReducedCosts(j) = "0"
22350                 End If
22360                 VarCell(j) = NameValue
22370             End If
22380             j = j + 1
22390         Wend

22400     End If
22410     Close #1
22420     ReadCBCModel = True
ExitSub:
          
End Function

Sub ReadCBCSensitivityData(SolutionFilePathName As String)
      'Reads the two files with the limits on the bounds of shadow prices and reduced costs

          Dim RangeFilePathName As String, Stuff(5) As String, index2 As Integer
          Dim Line As String, row As Integer, j As Integer, i As Integer
22430     ReDim IncreaseCon(NumRows)
22440     ReDim DecreaseCon(NumRows)
          
          'Find the ranges on the constraints
22450     RangeFilePathName = Replace(SolutionFilePathName, "modelsolution", "rhsranges")
22460     Open RangeFilePathName For Input As 2 ' supply path with filename
22470     Line Input #2, Line 'Dont want first line
22480     j = 1
22490     While Not EOF(2)
22500         Line Input #2, Line
22510         For i = 1 To 5
22520             index2 = InStr(Line, ",")
22530             Stuff(i) = left(Line, index2 - 1)
22540             If Stuff(i) = "1e-007" Then
22550                 Stuff(i) = "0"
22560             ElseIf InStr(Stuff(i), "E-16") Then
22570                 Stuff(i) = "0"
22580             End If
22590             Line = Mid(Line, index2 + 1)
22600         Next i
22610         IncreaseCon(j) = Stuff(3)
22620         DecreaseCon(j) = Stuff(5)
22630         j = j + 1
22640     Wend
22650     Close 2
          
22660     j = 1
          'Find the ranges on the variables
22670     RangeFilePathName = Replace(SolutionFilePathName, "modelsolution", "costranges")
22680     Open RangeFilePathName For Input As 2 ' supply path with filename
22690     Line Input #2, Line 'Dont want first line
22700     row = NumRows + 2
22710     While Not EOF(2)
22720         ReDim Preserve IncreaseVar(j)
22730         ReDim Preserve DecreaseVar(j)
22740         Line Input #2, Line
22750         For i = 1 To 5
22760             index2 = InStr(Line, ",")
22770             Stuff(i) = left(Line, index2 - 1)
22780             If Stuff(i) = "1e-007" Then
22790                 Stuff(i) = "0"
22800             ElseIf InStr(Stuff(i), "E-16") Then
22810                 Stuff(i) = "0"
22820             End If
22830             Line = Mid(Line, index2 + 1)
22840         Next i
22850         If ObjectiveSense = MaximiseObjective Then
22860             IncreaseVar(j) = Stuff(5)
22870             DecreaseVar(j) = Stuff(3)
22880         Else
22890             IncreaseVar(j) = Stuff(3)
22900             DecreaseVar(j) = Stuff(5)
22910         End If
22920         j = j + 1
22930     Wend
22940     Close 2
                    
End Sub

Sub WriteConstraintSensitivityTable(nameSheet As String, SolverString As String)

      'Writes out the sensitivity table on a new page (like the solver sensitivity report)
          
          Dim Column As Integer, row As Integer, i As Long
22950     Sheets(nameSheet).Cells(1, 1) = "OpenSolver Sensitivity Report - " & SolverString
22960     Sheets(nameSheet).Cells(2, 1) = "Worksheet: [" & ActiveWorkbook.Name & "] " & ActiveSheet.Name
22970     Sheets(nameSheet).Cells(3, 1) = "Report Created: " & Now()
          
22980     Column = 2
22990     row = 6
          
          Dim sheet As String, headings As Variant
          
23000     sheet = ActiveSheet.Name
23010     headings = Array("Cells", "Name", "Final Value", "Reduced Costs", "Objective Value", "Allowable Increase", "Allowable Decrease")
          'headings for the variable table
23020     Sheets(nameSheet).Cells(row - 1, Column - 1) = "Decision Variables"
23030     For i = 1 To UBound(headings)
23040         Sheets(nameSheet).Cells(row, Column + i - 1) = headings(i)
23050     Next i
23060     row = row + 1
          
          Dim numVars As Double
23070     numVars = UBound(ReducedCosts)
          'put the values into the variable table
23080     For i = 1 To numVars
23090         Sheets(nameSheet).Cells(row, Column) = VarCell(i)
23100         Sheets(nameSheet).Cells(row, Column + 2) = FinalVarValue(i)
23110         Sheets(nameSheet).Cells(row, Column + 3) = ReducedCosts(i)
23120         Sheets(nameSheet).Cells(row, Column + 4) = CostCoeffs(i)
23130         Sheets(nameSheet).Cells(row, Column + 5) = IncreaseVar(i)
23140         Sheets(nameSheet).Cells(row, Column + 6) = DecreaseVar(i)
23150         Sheets(nameSheet).Cells(row, Column + 1) = findName(sheet, VarCell(i))
23160         row = row + 1
23170     Next i
          
23180     row = row + 2
          
23190     headings(4) = "Shadow Price"
23200     headings(5) = "RHS Value"

          'Headings for constraint table
23210     Sheets(nameSheet).Cells(row - 1, Column - 1) = "Constraints"
23220     For i = 1 To UBound(headings)
23230         Sheets(nameSheet).Cells(row, Column + i - 1) = headings(i)
23240     Next i

23250     row = row + 1
          Dim constraint As Long, instance As Long
23260     constraint = 1

          'Values for constraint table
23270     For i = 1 To NumRows
23280         Sheets(nameSheet).Cells(row, Column + 2) = FinalValue(i)
23290         Sheets(nameSheet).Cells(row, Column + 3) = ShadowPrice(i)
23300         Sheets(nameSheet).Cells(row, Column + 4) = RHS(i)
23310         Sheets(nameSheet).Cells(row, Column + 5) = IncreaseCon(i)
23320         Sheets(nameSheet).Cells(row, Column + 6) = DecreaseCon(i)
              'This finds the range for cells of each constraint (similar to WriteConstraintListToSheet)
23330         GetConstraintFromRow i, constraint, instance  ' Which Excel constraint are we in, and which instance?
              Dim UnusedConstraint As Boolean
23340         UnusedConstraint = SparseA(i).Count = 0
              Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
23350         GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
23360         If Not RHSCellRange Is Nothing Then
23370             RHSstring = RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
23380         Else
23390             RHSstring = RemoveActiveSheetNameFromString(RHSstring)  ' Make this formula or value prettier
23400             RHSstring = Replace(RHSstring, "$", "")
23410         End If
              Dim s As String
23420         s = IIf(UnusedConstraint, "", "") & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & _
              SolverRelationAsString(Relation(i)) & RHSstring & IIf(UnusedConstraint, "", "")
              'Cell Range for each constraint
23430         Sheets(nameSheet).Cells(row, Column).value = s
              'Finds the nearest name for the constraint
23440         Sheets(nameSheet).Cells(row, Column + 1) = findName(sheet, LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False))
23450         row = row + 1
23460     Next i

          'Format the sensitivity table
23470     Call FormatSensitivityTable(nameSheet, row, numVars)
          
End Sub

Sub FormatSensitivityTable(nameSheet As String, row As Integer, numVars As Double)
          'Formats the sensitivity table on the new page with borders and bold writing
          
23480     Application.ScreenUpdating = False
          Dim sheet As String, startRow As String
23490     sheet = ActiveSheet.Name
23500     Sheets(nameSheet).Select
23510     startRow = 6
          
23520     Sheets(nameSheet).Cells.EntireColumn.AutoFit
23530     Columns("A:A").ColumnWidth = 5
23540     With Sheets(nameSheet).Range(Cells(2, 2), Cells(row, 8))
23550         .HorizontalAlignment = xlCenter
23560     End With
          
          'Create the borders for the constraint table
23570     Sheets(nameSheet).Range(Cells(startRow, 2), Cells(startRow + numVars, 8)).Select
23580     With Selection.Borders(xlEdgeLeft)
23590         .LineStyle = xlContinuous
23600         .ColorIndex = 0
23610         .TintAndShade = 0
23620         .Weight = xlMedium
23630     End With
23640     With Selection.Borders(xlEdgeTop)
23650         .LineStyle = xlContinuous
23660         .ColorIndex = 0
23670         .TintAndShade = 0
23680         .Weight = xlMedium
23690     End With
23700     With Selection.Borders(xlEdgeBottom)
23710         .LineStyle = xlContinuous
23720         .ColorIndex = 0
23730         .TintAndShade = 0
23740         .Weight = xlMedium
23750     End With
23760     With Selection.Borders(xlEdgeRight)
23770         .LineStyle = xlContinuous
23780         .ColorIndex = 0
23790         .TintAndShade = 0
23800         .Weight = xlMedium
23810     End With
23820     With Selection.Borders(xlInsideVertical)
23830         .LineStyle = xlContinuous
23840         .Weight = xlThin
23850     End With
          
          'Create the borders for the variable table
23860     Sheets(nameSheet).Range(Cells(numVars + startRow + 3, 2), Cells(row - 1, 8)).Select
23870     With Selection.Borders(xlEdgeLeft)
23880         .LineStyle = xlContinuous
23890         .ColorIndex = 0
23900         .TintAndShade = 0
23910         .Weight = xlMedium
23920     End With
23930     With Selection.Borders(xlEdgeTop)
23940         .LineStyle = xlContinuous
23950         .ColorIndex = 0
23960         .TintAndShade = 0
23970         .Weight = xlMedium
23980     End With
23990     With Selection.Borders(xlEdgeBottom)
24000         .LineStyle = xlContinuous
24010         .ColorIndex = 0
24020         .TintAndShade = 0
24030         .Weight = xlMedium
24040     End With
24050     With Selection.Borders(xlEdgeRight)
24060         .LineStyle = xlContinuous
24070         .ColorIndex = 0
24080         .TintAndShade = 0
24090         .Weight = xlMedium
24100     End With
24110     With Selection.Borders(xlInsideVertical)
24120         .LineStyle = xlContinuous
24130         .Weight = xlThin
24140     End With
          
          'Bold the constraint table headings and make them blue as well as put a border around them
24150     Sheets(nameSheet).Range(Cells(numVars + startRow + 3, 2), Cells(numVars + startRow + 3, 8)).Select
24160     With Selection.Borders(xlEdgeBottom)
24170         .LineStyle = xlContinuous
24180         .ColorIndex = 0
24190         .TintAndShade = 0
24200         .Weight = xlMedium
24210     End With
24220     With Selection.Font
24230         .Bold = True
24240         .ThemeColor = xlThemeColorLight2
24250     End With
          
          'Bold the variable table headings and make them blue as well as put a border around them
24260     Sheets(nameSheet).Range(Cells(startRow, 2), Cells(startRow, 8)).Select
24270     With Selection.Borders(xlEdgeBottom)
24280         .LineStyle = xlContinuous
24290         .ColorIndex = 0
24300         .TintAndShade = 0
24310         .Weight = xlMedium
24320     End With
24330     With Selection.Font
24340         .Bold = True
24350         .ThemeColor = xlThemeColorLight2
24360     End With
          
          'Bold the headings
24370     With Range("A:A").Font
24380         .Bold = True
24390     End With
          
24400     Cells(100, 100).Select
24410     Sheets(sheet).Select
24420     Application.ScreenUpdating = True
End Sub

Function findName(sheet As String, cell As String) As String
          'Finds the name of a constraint or variable by finding the nearest strings to the left and
          'above the cell and putting these together
          
          Dim NotFoundStringLeft As Boolean, NotFoundStringTop As Boolean
          Dim row As Integer, col As Integer, i As Integer, j As Integer
          Dim CellValue As String, LHSName As String, AboveName As String
          
24430     row = Sheets(sheet).Range(cell).row
24440     col = Sheets(sheet).Range(cell).Column
24450     i = col - 1
24460     j = row - 1
          
24470     NotFoundStringLeft = True
24480     NotFoundStringTop = True
          'Loop through to the left and above the cell to find the first non-numeric cell
24490     While (NotFoundStringLeft Or NotFoundStringTop) And (i > 0) And (j > 0)
              'Find the nearest name to the left of the variable or constraint if one exists
24500         If i > 0 And NotFoundStringLeft Then
24510             CellValue = Sheets(sheet).Cells(row, i)
                  'x = IsAmericanNumber(CellValue)
24520             If Not (IsNumeric(CellValue) Or (left(CellValue, 1) = "=") Or (CellValue = "")) Then
24530                 NotFoundStringLeft = False
24540             End If
24550             i = i - 1
24560         End If
              'Find the nearest name above the variable or constraint if it exists
24570         If j > 0 And NotFoundStringTop Then
24580             CellValue = Sheets(sheet).Cells(j, col)
24590             If Not (IsNumeric(CellValue) Or (left(CellValue, 1) = "=") Or (CellValue = "")) Then
24600                 NotFoundStringTop = False
24610             End If
24620             j = j - 1
24630         End If
24640     Wend
24650     LHSName = Cells(row, i + 1)
24660     AboveName = Cells(j + 1, col)
          
          'Put the names together
24670     If AboveName = "" Then
24680         findName = LHSName
24690     ElseIf LHSName = "" Then
24700         findName = AboveName
24710     Else
24720         findName = LHSName & " " & AboveName
24730     End If

End Function

'====================================================================
'********************NON-LINEAR**************************************
'====================================================================

Function updateVarOS(X As Variant)

          Dim i As Integer, numVars As Integer

          'set new variable values on sheet
24740     numVars = UBound(X)
24750     i = 1
          Dim AdjCell As Range
24760     For Each AdjCell In AdjustableCells
24770         AdjCell.Value2 = X(i, 1)
24780         i = i + 1
24790     Next AdjCell

End Function

Function getValuesOS() As Variant
          Dim y As Variant, i As Integer, j As Integer, k As Integer, conCountB As Integer, numCons As Variant
24800     numCons = getNumConstraintsOS()
24810     ReDim y(1 To numCons(0), 1 To 1)
          '====NOMAD only does minimise so need to change objective if it is max====
          'If objective sense is maximise then multiply by minus 1
24820     If ObjectiveSense = MaximiseObjective Then
24830         If ObjRange.Value2 <> 0 Then
24840             y(1, 1) = -1 * ObjRange.Value2 'objective value
24850         Else
24860             y(1, 1) = ObjRange.Value2
24870         End If
          'Else if objective sense is minimise leave it
24880     ElseIf ObjectiveSense = MinimiseObjective Then
24890         y(1, 1) = ObjRange.Value2
24900     ElseIf ObjectiveSense = TargetObjective Then
24910         y(1, 1) = Abs(ObjRange.Value2 - ObjectiveTargetValue)
24920     End If
24930     conCountB = 1 'keep a count of what constraint its up to including bounds
24940     k = 1 'keep a count of what constraint its up to not including bounds
          Dim row As Long, constraint As Long
24950     row = 1
          Dim CurrentLHSValues As Variant
          Dim CurrentRHSValues As Variant
24960     For constraint = 1 To NumConstraints
              ' Check to see what is different and add rows to sparsea
24970         If Not LHSRange(constraint) Is Nothing Then ' skip Binary and Integer constraints
                  ' Get current value(s) for LHS and RHS of this constraint off the sheet. LHS is always an array (even if 1x1)
24980             GetCurrentConstraintValues constraint, CurrentLHSValues, CurrentRHSValues
                  Dim coeff As Double
24990             If RHSType(constraint) = SolverInputType.MultiCellRange Then
25000                 For i = 1 To UBound(LHSOriginalValues(constraint), 1)
25010                     For j = 1 To UBound(LHSOriginalValues(constraint), 2)
25020                         If VariableBounds(conCountB) = False Then
25030                             If Relation(conCountB) = 1 Then 'less then or equal to
25040                                 If UBound(LHSOriginalValues(constraint), 1) = UBound(RHSOriginalValues(constraint), 1) Then
25050                                     y(k + 1, 1) = CurrentLHSValues(i, j) - CurrentRHSValues(i, j)  'constraint value
25060                                 Else
25070                                     y(k + 1, 1) = CurrentLHSValues(j, i) - CurrentRHSValues(i, j)  'constraint value
25080                                 End If
25090                             ElseIf Relation(conCountB) = 3 Then 'greater then or equal to
25100                                 If UBound(LHSOriginalValues(constraint), 1) = UBound(RHSOriginalValues(constraint), 1) Then
25110                                     y(k + 1, 1) = CurrentRHSValues(i, j) - CurrentLHSValues(i, j)
25120                                 Else
25130                                     y(k + 1, 1) = CurrentRHSValues(j, i) - CurrentLHSValues(i, j)
25140                                 End If
                                  '--------what to do in this case--------------
25150                             ElseIf Relation(conCountB) = 2 Then 'equal to
25160                                 If UBound(LHSOriginalValues(constraint), 1) = UBound(RHSOriginalValues(constraint), 1) Then
25170                                     y(k + 1, 1) = CurrentLHSValues(i, j) - CurrentRHSValues(i, j)
25180                                     y(k + 2, 1) = CurrentRHSValues(i, j) - CurrentLHSValues(i, j)
25190                                 Else
25200                                     y(k + 1, 1) = CurrentLHSValues(j, i) - CurrentRHSValues(i, j)
25210                                     y(k + 2, 1) = CurrentRHSValues(j, i) - CurrentLHSValues(i, j)
25220                                 End If
25230                                 k = k + 1
25240                             End If
25250                             k = k + 1
25260                         End If
25270                         conCountB = conCountB + 1
25280                     Next j
25290                 Next i
25300             Else
25310                 If VariableBounds(conCountB) = False Then
25320                     If Relation(conCountB) = RelationLE Then 'less then or equal to
25330                         y(k + 1, 1) = CurrentLHSValues(1, 1) - CurrentRHSValues 'constraint value
25340                     ElseIf Relation(conCountB) = RelationGE Then 'greater then or equal to
25350                         y(k + 1, 1) = CurrentRHSValues - CurrentLHSValues(1, 1)
                          '--------what to do in this case--------------
25360                     ElseIf Relation(conCountB) = RelationEQ Then 'equal to
25370                         y(k + 1, 1) = CurrentLHSValues(1, 1) - CurrentRHSValues
25380                         y(k + 2, 1) = CurrentRHSValues - CurrentLHSValues(1, 1)
25390                         k = k + 1
25400                     End If
25410                     k = k + 1
25420                 End If
25430                 conCountB = conCountB + 1
25440             End If
25450         End If
25460     Next constraint
          
          'Get back new objective and difference between LHS and RHS values
25470     getValuesOS = y
End Function

Function getNumVariablesOS() As Variant
25480     getNumVariablesOS = AdjustableCells.Count
End Function

Function getNumConstraintsOS() As Variant
          'The number of constraints is actually the number of Objectives + Number of Constraints
          'Note: Bounds do not count as constraints and equalities count as 2 constraints
          Dim i As Integer
          Dim X(0 To 1) As Double
25490     X(0) = 1
25500     For i = 1 To NumRows
25510         If VariableBounds(i) = False Then X(0) = X(0) + 1
25520         If Relation(i) = RelationEQ Then X(0) = X(0) + 1
25530     Next i

          'Number of objectives - NOMAD can do bi-objective
          'will take the first two values passed to it from getValuesOS as the objectives if doing biobj
          'Note: Currently OpenSolver can only do single objectives- will need to set up multi objectives yourself
25540     X(1) = 1 'number of objectives
25550     getNumConstraintsOS = X
End Function

Function getVariableDataOS() As Variant
          Dim numVars As Double
25560     numVars = AdjustableCells.Count
          Dim X() As Double
25570     ReDim X(0 To 4 * numVars - 1)
          Dim i As Long, j As Long
25580     For i = 0 To numVars - 1
25590         If AssumeNonNegativeVars Then
25600             X(2 * i) = 0
25610         Else
25620             X(2 * i) = -10000000000000#
25630         End If
25640         X(2 * i + 1) = 10000000000000#
25650     Next i
          
          Dim BoundedVariables As Range, row As Long, var As Long, constraint As Long, instance As Long
25660     constraint = 1
25670     ReDim VariableBounds(NumRows)
25680     For row = 1 To NumRows
25690         VariableBounds(row) = False
25700         If (Relation(row) = RelationGE Or Relation(row) = RelationLE) And SparseA(row).Count = 1 Then
              ' This is a constraint apply a lower bound to one variable. Check that the LHS did directly specify the decision variable.
25710             var = SparseA(row).Index(1)
25720             GetConstraintFromRow row, constraint, instance
25730             If RHSType(constraint) <> Formula Then
25740                 GetArrayPosnFromConstraintInstance constraint, instance, i, j
25750                 If LHSRange(constraint)(i, j).Address(RowAbsolute:=False, ColumnAbsolute:=False) = VarNames(var) _
                          And LHSRange(constraint).Worksheet.Name = AdjustableCells.Worksheet.Name Then ' AJM Must be on the same sheet
                          ' This is a direct lower bound on this variable
25760                     If Relation(row) = 1 Then
25770                         If RHSType(constraint) = MultiCellRange Then
25780                             X(2 * var - 1) = RHSOriginalValues(constraint)(i, j)
25790                         Else
25800                             X(2 * var - 1) = RHSOriginalValues(constraint)
25810                         End If
25820                     ElseIf Relation(row) = 3 Then
25830                         If RHSType(constraint) = MultiCellRange Then
25840                             X(2 * var - 2) = RHSOriginalValues(constraint)(i, j)
25850                         Else
25860                             X(2 * var - 2) = RHSOriginalValues(constraint)
25870                         End If
25880                     End If
25890                     VariableBounds(row) = True
25900                 End If
25910             End If
25920         End If
25930     Next row

          'Get the starting point
          'Takes the points on the sheet and forces them between the bounds
25940     For j = 0 To numVars - 1
25950         If StartingPosition(j + 1) < X(2 * j) Then
25960             X(j + 2 * numVars) = X(2 * j)
25970         ElseIf StartingPosition(j + 1) > X(2 * j + 1) Then
25980             X(j + 2 * numVars) = X(2 * j + 1)
25990         Else
26000             X(j + 2 * numVars) = StartingPosition(j + 1)
26010         End If
26020     Next j
          
          'Get the variable type(real, int or bin)
26030     For i = 1 To numVars
          'initialise all variables as continuous
26040         X(i - 1 + 3 * numVars) = 1
26050     Next i
          Dim counter As Integer, types As Variant, c As Range
26060     counter = 2
26070     For Each types In Array(IntegerCellsRange, BinaryCellsRange)
26080         If Not types Is Nothing Then
26090             For Each c In types
26100                 For i = 1 To numVars
26110                     If VarNames(i) = c.Address(RowAbsolute:=False, ColumnAbsolute:=False) Then
26120                         X(i - 1 + 3 * numVars) = counter
26130                         If Not SolveRelaxedModel Then
                                  'Make bounds on integer and binary constraints integer
26140                             If X(2 * i - 2) > 0 Then
26150                                 X(2 * i - 2) = Application.WorksheetFunction.RoundUp(X(2 * i - 2), 0)
26160                             Else
26170                                 X(2 * i - 2) = Application.WorksheetFunction.RoundDown(X(2 * i - 2), 0)
26180                             End If
26190                             If X(2 * i - 1) > 0 Then
26200                                 X(2 * i - 1) = Application.WorksheetFunction.RoundDown(X(2 * i - 1), 0)
26210                             Else
26220                                 X(2 * i - 1) = Application.WorksheetFunction.RoundUp(X(2 * i - 1), 0)
26230                             End If
                                  'Make starting positions on integer and binary constraints integer
26240                             If X(i - 1 + 2 * numVars) < X(2 * i - 2) Then
26250                                 X(i - 1 + 2 * numVars) = X(2 * i - 2)
26260                             ElseIf X(i - 1 + 2 * numVars) > X(2 * i - 1) Then
26270                                 X(i - 1 + 2 * numVars) = X(2 * i - 1)
26280                             Else
26290                                 X(i - 1 + 2 * numVars) = Round(X(i - 1 + 2 * numVars))
26300                             End If
26310                         End If
26320                     End If
26330                 Next i
26340             Next c
26350         End If
26360         counter = counter + 1
26370     Next types
          
26380     getVariableDataOS = X
End Function

Function getOptionDataOS() As Variant
          Dim X() As Double
26390     ReDim X(0 To 2)
          Dim SolveOptions As SolveOptionsType, ErrorString As String
26400     GetSolveOptions sheetName, SolveOptions, ErrorString
26410     X(0) = SolveOptions.MaxIterations
26420     X(1) = SolveOptions.maxTime
26430     X(2) = SolveOptions.Precision
26440     getOptionDataOS = X
End Function

