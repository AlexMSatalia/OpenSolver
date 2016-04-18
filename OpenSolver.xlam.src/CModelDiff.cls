VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CModelDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private s As COpenSolver

' =======================================
' Model building functions
' =======================================

Public Function ProcessSolverModel(OpenSolver As COpenSolver, LinearityOffset As Double, ShouldCheckLinearity As Boolean)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Set s = OpenSolver
          
          If Not BuildConstantTerms() Then GoTo ExitFunction
          If Not BuildVariableTerms(LinearityOffset) Then GoTo ExitFunction
          
          If Not QuickFeasibilityCheck() Then GoTo ExitFunction
          
          If ShouldCheckLinearity Then
              DoQuickLinearityCheck
          End If
          
          s.AdjustableCells.Value2 = 0
          ProcessSolverModel = True

ExitFunction:
          Application.StatusBar = False
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function
          
ErrorHandler:
          If Not ReportError("CModelDiff", "ProcessSolverModel") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Private Function BuildConstantTerms() As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          'Create the rhs vector by resetting cells to zero
          s.AdjustableCells.Value2 = 0
          ' Recalculate the spreadsheet
          If Not ForceCalculate("Warning: The worksheet calculation did not complete, and so the model may not be generated correctly. Would you like to retry?") Then GoTo ExitFunction

          ' TODO Handle errors in the original spreadsheet values
          Dim ZeroedLHSValues As Variant, ZeroedRHSValues As Variant
          Dim row As Long, constraint As Long, i As Long, j As Long
1320      row = 1
1321      For constraint = 1 To s.NumConstraints
1322          If Not s.LHSRange(constraint) Is Nothing Then
                  ' Get current value(s) for LHS and RHS of this constraint off the sheet. LHS is always an array (even if 1x1)
                  s.GetCurrentConstraintValues constraint, ZeroedLHSValues, ZeroedRHSValues
1323              If s.RHSType(constraint) = SolverInputType.MultiCellRange Then 'does the constraint refer to more than one cell
1324                  For i = 1 To UBound(ZeroedLHSValues, 1)
1325                      For j = 1 To UBound(ZeroedLHSValues, 2)
1326                          If UBound(ZeroedLHSValues, 1) = UBound(ZeroedRHSValues, 1) Then
1327                              s.RHS(row) = ZeroedRHSValues(i, j) - ZeroedLHSValues(i, j)
1328                          Else
1329                              s.RHS(row) = ZeroedRHSValues(j, i) - ZeroedLHSValues(i, j)
1330                          End If
1331                          row = row + 1
1332                      Next j
1333                  Next i
1334              Else
1335                  For i = 1 To UBound(ZeroedLHSValues, 1)
1336                      For j = 1 To UBound(ZeroedLHSValues, 2)
1337                          s.RHS(row) = ZeroedRHSValues - ZeroedLHSValues(i, j)
1338                          row = row + 1
1339                      Next j
1340                  Next i
1341              End If
1342          End If
1343      Next constraint

          If Not s.ObjRange Is Nothing Then
              s.ObjectiveFunctionConstant = s.GetCurrentObjectiveValue()
              ' Add constants to Target
              If s.ObjectiveSense = TargetObjective Then
                  ' Target value needs to be incremented by any constants in the objective
                  s.ObjectiveTargetValue = s.ObjectiveTargetValue - s.ObjectiveFunctionConstant
              End If
          End If
          
          BuildConstantTerms = True

ExitFunction:
          Application.StatusBar = False
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function
          
ErrorHandler:
          If Not ReportError("CModelDiff", "BuildConstantTerms") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Private Function BuildVariableTerms(LinearityOffset As Double, Optional BinaryOffset As Double = 0) As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          ' Initialize
          Dim row As Long
1182      For row = 1 To s.NumRows
1183          Set s.SparseA(row) = New CIndexedCoeffs
1184      Next row
          Set s.CostCoeffs = New CIndexedCoeffs

          s.AdjustableCells.Value2 = LinearityOffset
          If Not s.BinaryCellsRange Is Nothing Then s.BinaryCellsRange.Value2 = BinaryOffset
          
          If Not ForceCalculate("Warning: The worksheet calculation did not complete, and so the model may not be generated correctly. Would you like to retry?") Then GoTo ExitFunction

          Dim LHSOriginalValues() As Variant, RHSOriginalValues() As Variant
          ReDim LHSOriginalValues(s.NumConstraints)
          ReDim RHSOriginalValues(s.NumConstraints)
          Dim constraint As Long
          For constraint = 1 To s.NumConstraints
              If Not s.LHSRange(constraint) Is Nothing Then
                  s.GetCurrentConstraintValues constraint, LHSOriginalValues(constraint), RHSOriginalValues(constraint)
              End If
          Next constraint
          
          Dim ObjPresent As Boolean, OriginalObjectiveValue As Double
          ObjPresent = Not s.ObjRange Is Nothing
          If ObjPresent Then OriginalObjectiveValue = s.GetCurrentObjectiveValue()
          
          ' Increment each decision variable by one to see what constraints are dependent on it and by what amount
          Dim var As Long, AdjCell As Range, coeff As Double
1279      var = 0
1280      For Each AdjCell In s.AdjustableCells
              ' Count how many variables we've processed, let the user know that the program hasn't crashed!
1281          var = var + 1
1282          UpdateStatusBar "OpenSolver: Setting Up Problem... " & var & "/" & s.NumVars & " vars, " & s.NumRows & " rows."
              
              ' We start binaries at 0 to test 0 and 1, and all others at the specified offset value
              Dim BaseValue As Double, NewValue As Double
              If ProperIntersect(AdjCell, s.BinaryCellsRange) Is Nothing Then
                  BaseValue = LinearityOffset
                  NewValue = LinearityOffset + 1
              Else
                  BaseValue = BinaryOffset
                  NewValue = 1 - BinaryOffset
              End If
1283          AdjCell.Value2 = NewValue

1284          If Not ForceCalculate("Warning: The worksheet calculation did not complete, and so the model may not be generated correctly. Would you like to retry?") Then GoTo ExitFunction

              ' Easy part first - the objective function value change
1285          If ObjPresent Then
                  coeff = s.GetCurrentObjectiveValue - OriginalObjectiveValue
1286              If Abs(coeff) > EPSILON Then
                      s.CostCoeffs.Add var, coeff
                  End If
1287          End If

              ' Now the hard (slow) bit - the constraint changes
1288          row = 1
1289          For constraint = 1 To s.NumConstraints

                  ' Check to see what is different and add rows to SparseA
1290              If Not s.LHSRange(constraint) Is Nothing Then ' skip Binary and Integer constraints

                      Dim OriginalLHSValues As Variant
                      Dim OriginalRHSValues As Variant
1291                  OriginalLHSValues = LHSOriginalValues(constraint)
1292                  OriginalRHSValues = RHSOriginalValues(constraint)

                      ' Get current value(s) for LHS and RHS of this constraint off the sheet. LHS is always an array (even if 1x1)
                      Dim CurrentLHSValues As Variant
                      Dim CurrentRHSValues As Variant
                      s.GetCurrentConstraintValues constraint, CurrentLHSValues, CurrentRHSValues

                      Dim i As Long, j As Long
1294                  For i = 1 To UBound(OriginalLHSValues, 1) ' number of rows
1295                      For j = 1 To UBound(OriginalLHSValues, 2) ' number of columns
1296                          Select Case s.RHSType(constraint)
                              Case SolverInputType.MultiCellRange
                                  '-----------------------------------------------------------
                                  'Making it work for column LHS with row RHS and vice versa
                                  '-----------------------------------------------------------
1297                              If UBound(OriginalLHSValues, 1) = UBound(OriginalRHSValues, 1) Then
1298                                  coeff = (CurrentLHSValues(i, j) - OriginalLHSValues(i, j)) - (CurrentRHSValues(i, j) - OriginalRHSValues(i, j))
1299                              Else
1300                                  coeff = (CurrentLHSValues(i, j) - OriginalLHSValues(i, j)) - (CurrentRHSValues(j, i) - OriginalRHSValues(j, i))
1301                              End If
1302                          Case SolverInputType.Formula, SolverInputType.SingleCellRange
1303                              coeff = (CurrentLHSValues(i, j) - OriginalLHSValues(i, j)) - (CurrentRHSValues - OriginalRHSValues)
1304                          Case SolverInputType.constant
1305                              coeff = (CurrentLHSValues(i, j) - OriginalLHSValues(i, j))
1306                          End Select
                              ' Ignore zero (or near-zero) coefficients
1307                          If Abs(coeff) > EPSILON Then
1308                              s.SparseA(row).Add var, coeff
1309                          End If
1310                          row = row + 1
1311                      Next j
1312                  Next i
1313              End If
1314          Next constraint

              AdjCell.Value2 = BaseValue
1316      Next AdjCell
          
1344      BuildVariableTerms = True
          
ExitFunction:
          Application.StatusBar = False
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("COpenSolver", "BuildVariableTerms") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function QuickFeasibilityCheck() As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
    
          ' Check if any empty constraints are feasible
          Dim row As Long
          For row = 1 To s.NumRows
              If s.SparseA(row).Count = 0 Then
                  ' We have a constraint that does not vary with the decision variables; check it is satisfied
                  Dim constraint As Long, instance As Long
                  constraint = s.RowToConstraint(row)
                  instance = s.GetConstraintInstance(row, constraint)
                  
                  Dim rel As RelationConsts, value As Double
                  rel = s.Relation(constraint)
                  value = s.RHS(row)
                  
                  If (rel = RelationEQ And Abs(value) > EPSILON) Or _
                     (rel = RelationGE And value > EPSILON) Or _
                     (rel = RelationLE And value < -EPSILON) Then
                      
                      Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
1894                  s.GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
                      ' If the RHS is a range, we show its address; if not, RHSString contains the RHS's constant or formula
1895                  If Not RHSCellRange Is Nothing Then RHSstring = RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)

                      Dim LHSValue As Double, RHSValue As Double
                      LHSValue = LHSCellRange.Value2
                      RHSValue = s.RHS(row) + LHSValue
                      
1896                  s.SolveStatus = OpenSolverResult.Infeasible
1897                  s.SolveStatusString = "Infeasible"
1898                  s.SolveStatusComment = "The model contains a constraint '" & s.ConstraintSummary(constraint) & "' for which instance " & _
                                             instance & " does not depend on the decision variables and is not satisfied." & vbCrLf & _
                                             "Constraint specifies: " & _
                                             "LHS=" & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "=" & LHSValue & _
                                             " " & RelationEnumToString(rel) & " " & _
                                             "RHS=" & StripWorksheetNameAndDollars(RHSstring, s.sheet) & "=" & RHSValue
                      GoTo ExitFunction
                  End If
              End If
          Next row
          
          ' Check objective if it is a target and all zeros
          If s.ObjectiveSense = TargetObjective Then
              If s.CostCoeffs.Count = 0 And _
                 Abs(s.ObjectiveTargetValue - s.ObjectiveFunctionConstant) > EPSILON Then
                  s.SolveStatus = OpenSolverResult.Infeasible
                  s.SolveStatusString = "Infeasible"
                  s.SolveStatusComment = "The objective cell does not depend on the adjustable cells " & _
                                         "and has a value different to the specified target value. " & _
                                         "This means the target value cannot be attained."
                  GoTo ExitFunction
              End If
          End If
          
          QuickFeasibilityCheck = True

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("CModelDiff", "QuickFeasibilityCheck") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

' =========================================
' Linearity check functions
' =========================================

' Highlight all constraints (and the objective) that are non-linear using our standard model highlighting, but showing only individual cells, not ranges
Sub HighlightNonLinearities(RowIsNonLinear() As Boolean, ObjectiveIsNonLinear As Boolean)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

2028      If SheetHasOpenSolverHighlighting(s.sheet) Then HideSolverModel s.sheet
2031      DeleteOpenSolverShapes s.sheet

2032      InitialiseHighlighting

          Dim constraint As Long, row As Long, instance As Long
2034      For row = 1 To s.NumRows
2035          If RowIsNonLinear(row) Then
                  constraint = s.RowToConstraint(row)
2036              instance = s.GetConstraintInstance(row, constraint)

                  Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
2037              s.GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring

2038              RHSstring = StripWorksheetNameAndDollars(RHSstring, LHSCellRange.Worksheet) ' Strip any worksheet name and $'s from the RHS (useful if it is a formula)
2040              HighlightConstraint LHSCellRange.Worksheet, LHSCellRange, RHSCellRange, RHSstring, s.Relation(constraint), 0  ' Show either a value or a formula
2041          End If
2042      Next row

2043      If ObjectiveIsNonLinear Then
2047          AddObjectiveHighlighting s.ObjRange, s.ObjectiveSense, s.ObjectiveTargetValue
2048      End If

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("CModelDiff", "HighlightNonLinearities") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Private Sub ExtractModelInfo(ByRef value() As CIndexedCoeffs, ByRef ObjectiveCoeffs As CIndexedCoeffs)
          If s.NumRows > 0 Then ReDim Preserve value(s.NumRows) As CIndexedCoeffs

          Dim row As Long
2055      For row = 1 To s.NumRows
2056          Set value(row) = s.SparseA(row).Clone
2057      Next row

          Set ObjectiveCoeffs = s.CostCoeffs.Clone
End Sub

Private Function RatioTest(Value1 As Double, Value2 As Double, Optional Threshold As Double = EPSILON) As Boolean
          RatioTest = Abs(Value1 - Value2) / (1 + Abs(Value1)) > Threshold
End Function

Private Sub DoFullLinearityCheck()
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim InteractiveStatus As Boolean
          InteractiveStatus = Application.Interactive
          
          Dim NonLinearInformation As String
2051      NonLinearInformation = ""
          
          'Build each matrix where the decision variables start at the base linearity offset (ValueBase()), one (ValueOne()) and ten (ValueTen())
          Dim ValueBase() As CIndexedCoeffs, ObjectiveCoeffsBase As CIndexedCoeffs
          ExtractModelInfo ValueBase, ObjectiveCoeffsBase
          
          Dim ValueOne() As CIndexedCoeffs, ObjectiveCoeffsOne As CIndexedCoeffs
2062      BuildVariableTerms LinearityOffset:=1
          ExtractModelInfo ValueOne, ObjectiveCoeffsOne
              
          Dim ValueTen() As CIndexedCoeffs, ObjectiveCoeffsTen As CIndexedCoeffs
          BuildVariableTerms LinearityOffset:=10, BinaryOffset:=1
          ExtractModelInfo ValueTen, ObjectiveCoeffsTen
          
          Dim RowIsNonLinear() As Boolean
2054      If s.NumRows > 0 Then ReDim RowIsNonLinear(s.NumRows) As Boolean

          Dim row As Long
          For row = 1 To s.NumRows
2082          RowIsNonLinear(row) = False
              
              Dim NonLinearConVars As String
              NonLinearConVars = CompareSparseVectors(ValueBase(row), ValueOne(row), ValueTen(row))
              If Len(NonLinearConVars) > 0 Then
                  Dim constraint As Long, instance As Long
                  constraint = s.RowToConstraint(row)
                  instance = s.GetConstraintInstance(row, constraint)
                  
                  If s.LHSType(constraint) = SolverInputType.SingleCellRange Then
2160                  NonLinearInformation = NonLinearInformation & vbNewLine & "In the constraint: " & s.ConstraintSummary(constraint) & "," & vbNewLine & "  the model appears to be non-linear in the decision variables: "
2161              Else
                      NonLinearInformation = NonLinearInformation & vbNewLine & "In instance " & instance & " of the constraint: " & s.ConstraintSummary(constraint) & "," & vbNewLine & "  the model appears to be non-linear in the following decision variables: "
2163              End If
                  NonLinearInformation = NonLinearInformation & NonLinearConVars
2101              RowIsNonLinear(row) = True
              End If
2105      Next row

          Dim ObjectiveIsNonLinear As Boolean, NonLinearObjVars As String
2109      ObjectiveIsNonLinear = False
          NonLinearObjVars = CompareSparseVectors(ObjectiveCoeffsBase, ObjectiveCoeffsOne, ObjectiveCoeffsTen)
2112      If Len(NonLinearObjVars) > 0 Then
2113          ObjectiveIsNonLinear = True
2114          NonLinearInformation = NonLinearInformation & vbNewLine & vbNewLine & _
                                     "The objective function is nonlinear in the following variables: " & _
                                     NonLinearObjVars
          End If

          NonLinearInformation = TrimBlankLines(NonLinearInformation)
          If NonLinearInformation = "" Then
              NonLinearInformation = "There have been no instances of nonlinearity found in this model. Some models can generate warnings of non-linearity " & _
                                     "because of numerical errors that accumulate in the spreadsheet. OpenSolver's non-linearity check can be disabled under OpenSolver's " & _
                                     "Options settings."
          End If
          
          'display dialog to user
          Dim frmNonlinear As FNonlinear
          Set frmNonlinear = New FNonlinear
          frmNonlinear.SetLinearityResult NonLinearInformation, False
          
          Application.Interactive = True
2139      frmNonlinear.Show
          Application.Interactive = InteractiveStatus
          
2140      If frmNonlinear.chkHighlight.value = True Then
2141          HighlightNonLinearities RowIsNonLinear, ObjectiveIsNonLinear
2142      End If
          Unload frmNonlinear

ExitSub:
          Application.Interactive = InteractiveStatus
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("CModelDiff", "DoFullLinearityCheck") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Function TestExistenceOfEntry(ValueBase As CIndexedCoeffs, ValueOne As CIndexedCoeffs, ValueTen As CIndexedCoeffs, i As Long) As Boolean
          'Check if this index exists in all three models
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

2144      If ValueBase.Index(i) <> 0 And ValueOne.Index(i) <> 0 And ValueTen.Index(i) <> 0 Then
2145          TestExistenceOfEntry = True
2146      End If

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
2147      If Err.Number = 9 Then
2148          TestExistenceOfEntry = False
              Resume ExitFunction
2149      End If

          If Not ReportError("CModelDiff", "TestExistenceOfEntry") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

'Return the ith entry from any one of these arrays; it may not exist in all of them
Function GetEntry(ValueBase As CIndexedCoeffs, ValueOne As CIndexedCoeffs, ValueTen As CIndexedCoeffs, i As Long) As Long
2150      On Error Resume Next
2151      If i <= ValueBase.Count Then
2152          GetEntry = ValueBase.Index(i)
2153      ElseIf i <= ValueOne.Count Then
2154          GetEntry = ValueOne.Index(i)
2155      ElseIf i <= ValueTen.Count Then
2156          GetEntry = ValueTen.Index(i)
2157      End If
End Function

Function CompareSparseVectors(v1 As CIndexedCoeffs, v2 As CIndexedCoeffs, v3 As CIndexedCoeffs) As String
' Returns a collection of all variable indices with different coefficients in the vectors
          Dim NumEntries As Long
          NumEntries = Max(v1.Count, v2.Count, v3.Count)
          
          Dim VarString As String
          
          Dim i As Long
          For i = 1 To NumEntries
              Dim VarIndex As Long
              VarIndex = 0
                  
              If TestExistenceOfEntry(v1, v2, v3, i) Then
                  'do a ratio test
                  If RatioTest(v1.Coefficient(i), v2.Coefficient(i)) Or _
                     RatioTest(v1.Coefficient(i), v3.Coefficient(i)) Then
                      VarIndex = v1.Index(i)
                  End If
              Else
                  VarIndex = GetEntry(v1, v2, v3, i)
              End If
              
              If VarIndex > 0 Then
                  VarString = VarString & ", " & s.VarName(VarIndex)
              End If
          Next i
          
          CompareSparseVectors = Mid(VarString, 3)  ' Trim leading comma
End Function

Sub DoQuickLinearityCheck()
' Returns false if a full check was performed by the user, meaning the model result is no longer valid.
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim InteractiveStatus As Boolean
          InteractiveStatus = Application.Interactive
          
          ' Do a better check for binary non-linearity - build from 0, check at 1
          If Not s.BinaryCellsRange Is Nothing Then
              s.BinaryCellsRange.Value2 = 1
          End If
          
2181      If Not ForceCalculate("Warning: The worksheet calculation did not complete during the linearity test, and so the test may not be correct. Would you like to retry?") Then GoTo ExitSub
          
          ' Get all the decision variable values off the sheet
          Dim DecisionVariableValues() As Double, var As Long, c As Range
2184      ReDim DecisionVariableValues(s.NumVars)
          var = 1
          For Each c In s.AdjustableCells
927           DecisionVariableValues(var) = c.Value2
              var = var + 1
928       Next c
          
          Dim NonLinearInfo As String
2186      NonLinearInfo = ""

          Dim RowIsNonLinear() As Boolean
2180      If s.NumRows > 0 Then ReDim RowIsNonLinear(s.NumRows) As Boolean
          
          Dim constraint As Long, row As Long
2187      row = 1
2188      For constraint = 1 To s.NumConstraints
2189          If Not s.LHSRange(constraint) Is Nothing Then
                  ' Get current value(s) for LHS and RHS of this constraint off the sheet
                  Dim CurrentLHSValues As Variant, CurrentRHSValues As Variant
2190              s.GetCurrentConstraintValues constraint, CurrentLHSValues, CurrentRHSValues
                  
                  Dim instance As Long, i As Long, j As Long
2194              instance = 0
2195              For i = 1 To UBound(CurrentLHSValues, 1)
2196                  For j = 1 To UBound(CurrentLHSValues, 2)
2197                      instance = instance + 1
                          
                          Dim SolutionValueLHS As Double, SolutionValueRHS As Double, SolutionValue As Double
2198                      SolutionValueLHS = CurrentLHSValues(i, j)
2199                      If s.RHSType(constraint) = SolverInputType.MultiCellRange Then
                              ' Check whether the LHS and RHS are parallel or perpendicular
2200                          If UBound(CurrentLHSValues, 1) = UBound(CurrentRHSValues, 1) Then
2201                              SolutionValueRHS = CurrentRHSValues(i, j)
2202                          Else
2203                              SolutionValueRHS = CurrentRHSValues(j, i)
2204                          End If
                          Else
                              SolutionValueRHS = CurrentRHSValues
2205                      End If
2206                      SolutionValue = SolutionValueLHS - SolutionValueRHS
                      
                          ' Find out what we expect the value to be from Ax = b. We track the maximum value we encounter during the calculation
                          ' so that we have some idea of the errors we might expect
                          Dim ExpectedValue As Double, MaxValueInCalculation As Double
2208                      ExpectedValue = s.SparseA(row).Evaluate_RecordPrecision(DecisionVariableValues, MaxValueInCalculation) - s.RHS(row)
2209                      MaxValueInCalculation = Max(MaxValueInCalculation, Abs(s.RHS(row)))
          
2210                      If RatioTest(ExpectedValue, SolutionValue, Threshold:=Max(EPSILON, EPSILON * MaxValueInCalculation)) Then
                              Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
2214                          s.GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
                              ' If the RHS is a range, we show its address; if not, RHSString contains the RHS's constant or formula
2215                          If Not RHSCellRange Is Nothing Then RHSstring = RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
                              
                              If NonLinearInfo = "" Then NonLinearInfo = "The following constraint(s) do not appear to be linear: "
2216                          NonLinearInfo = NonLinearInfo & vbNewLine & s.ConstraintSummary(constraint) & ": instance " & instance & _
                                              ", LHS=" & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & _
                                              ", RHS=" & RHSstring & _
                                              ", " & ExpectedValue & "<>" & SolutionValue
                              
2217                          RowIsNonLinear(row) = True
2221                      End If
2222                      row = row + 1
2223                  Next j
2224              Next i
2225          End If
2226      Next constraint

          'Check objective function for linearity
          Dim CalculatedObjValue As Double, ObservedObjValue As Double, ObjectiveIsNonLinear As Boolean
          If s.ObjRange Is Nothing Then
              ObservedObjValue = 0
          Else
2233          ObservedObjValue = s.GetCurrentObjectiveValue()
          End If
2236      CalculatedObjValue = s.CostCoeffs.Evaluate_RecordPrecision(DecisionVariableValues, MaxValueInCalculation) + s.ObjectiveFunctionConstant
          MaxValueInCalculation = Max(MaxValueInCalculation, Abs(s.ObjectiveFunctionConstant))

2237      ObjectiveIsNonLinear = RatioTest(CalculatedObjValue, ObservedObjValue)
2238      If ObjectiveIsNonLinear Then
2239         NonLinearInfo = "The objective function is not linear." & vbNewLine & vbNewLine & NonLinearInfo
2240      End If
          
          'Set the userform up and display any information on nonlinear constraints
2241      If Len(NonLinearInfo) > 0 Then
2242          s.SolveStatus = NotLinear
2243          If Not s.MinimiseUserInteraction Then
                  NonLinearInfo = "WARNING : " & vbNewLine & TrimBlankLines(NonLinearInfo)
                  
                  Dim frmNonlinear As FNonlinear
                  Set frmNonlinear = New FNonlinear
                  frmNonlinear.SetLinearityResult NonLinearInfo, True
                  
                  Application.Interactive = True
2263              frmNonlinear.Show
                  Application.Interactive = InteractiveStatus
              
                  'Show the nonlinear constraints if requested
2264              If frmNonlinear.chkHighlight.value = True Then
2265                  HighlightNonLinearities RowIsNonLinear, ObjectiveIsNonLinear
2266              End If
                  ' Conduct full linearity check if requested
2267              If frmNonlinear.chkFullCheck.value = True Then DoFullLinearityCheck

                  Unload frmNonlinear
2271          End If
2272      End If

ExitSub:
          Application.Interactive = InteractiveStatus
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("CModelDiff", "DoQuickLinearityCheck") Then Resume
          RaiseError = True
          GoTo ExitSub
          
End Sub

