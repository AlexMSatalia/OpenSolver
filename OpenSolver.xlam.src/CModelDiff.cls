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
          
          ' Initialize
          Dim row As Long
1182      For row = 1 To s.NumRows
1183          Set s.SparseA(row) = New CIndexedCoeffs
1184      Next row

          If Not BuildConstantTerms() Then GoTo ExitFunction
          If Not BuildVariableTerms(LinearityOffset) Then GoTo ExitFunction
          
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
              On Error GoTo NotLinear
              s.ObjectiveFunctionConstant = s.ObjRange.Value2
              On Error GoTo ErrorHandler
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
          
NotLinear:
          Err.Raise OpenSolver_ModelError, Description:="The objective cell appears to contain an error. This could have occurred if there is a divide by zero error or if you have used the wrong function (eg #DIV/0! or #VALUE!). Please fix this, and try again."
End Function

Private Function BuildVariableTerms(LinearityOffset As Double) As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          s.AdjustableCells.Value2 = LinearityOffset
          If Not ForceCalculate("Warning: The worksheet calculation did not complete, and so the model may not be generated correctly. Would you like to retry?") Then GoTo ExitFunction

          Dim LHSOriginalValues() As Variant, RHSOriginalValues() As Variant
          ReDim LHSOriginalValues(s.NumConstraints)
          ReDim RHSOriginalValues(s.NumConstraints)
          Dim constraint As Long
          For constraint = 1 To s.NumConstraints
              If Not s.LHSRange(constraint) Is Nothing Then
                  s.GetCurrentConstraintValues constraint, LHSOriginalValues(constraint), RHSOriginalValues(constraint)
                  ' Check that all our values are in fact numbers
                  Dim ErrorMsg As String
1171              If Not s.CheckConstraintValuesAreNumeric(constraint, LHSOriginalValues(constraint), RHSOriginalValues(constraint), ErrorMsg) Then
1173                  Err.Raise Number:=OpenSolver_BuildError, Description:=ErrorMsg
1174              End If
              End If
          Next constraint
          
          Dim ObjPresent As Boolean, OriginalObjectiveValue As Double
          ObjPresent = Not s.ObjRange Is Nothing
          If ObjPresent Then OriginalObjectiveValue = s.ObjRange.Value2
          
          ' Increment each decision variable by one to see what constraints are dependent on it and by what amount
          Dim var As Long, AdjCell As Range
1279      var = 0
1280      For Each AdjCell In s.AdjustableCells
              ' Count how many variables we've processed, let the user know that the program hasn't crashed!
1281          var = var + 1
1282          UpdateStatusBar "OpenSolver: Setting Up Problem... " & var & "/" & s.NumVars & " vars, " & s.NumRows & " rows."

1283          AdjCell.Value2 = AdjCell.Value2 + 1
1284          If Not ForceCalculate("Warning: The worksheet calculation did not complete, and so the model may not be generated correctly. Would you like to retry?") Then GoTo ExitFunction

              ' Easy part first - the objective function value change
1285          If ObjPresent Then
1286              s.CostCoeffs(var) = s.ObjRange.Value2 - OriginalObjectiveValue
1287          End If

              ' Now the hard (slow) bit - the constraint changes
              Dim row As Long
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
                              Dim coeff As Double
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
1315          AdjCell.Value2 = AdjCell.Value2 - 1
1316      Next AdjCell

          ' Now we check if any empty constraints are feasible
          For row = 1 To s.NumRows
              If s.SparseA(row).Count = 0 Then
                  ' We have a constraint that does not vary with the decision variables; check it is satisfied
                  Dim instance As Long
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

' =========================================
' Linearity check functions
' =========================================

Function CalcObjFnValue(DecisionVariableValues() As Double) As Double
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim sum As Double, var As Long
930       sum = s.ObjectiveFunctionConstant ' The value found for the obj function when all decision variables are 0
931       For var = 1 To s.NumVars
932           sum = sum + (s.CostCoeffs(var) * DecisionVariableValues(var))
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

Private Sub ExtractModelInfo(ByRef value() As CIndexedCoeffs, ByRef ObjectiveCoeffs() As Double)
          If s.NumRows > 0 Then ReDim Preserve value(s.NumRows) As CIndexedCoeffs
2053      ReDim ObjectiveCoeffs(s.NumVars) As Double

          Dim row As Long
2055      For row = 1 To s.NumRows
2056          Set value(row) = s.SparseA(row).Clone
2057      Next row

          Dim i As Long
2058      For i = 1 To s.NumVars
2059          ObjectiveCoeffs(i) = s.CostCoeffs(i)
2060      Next i
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
          Dim ValueBase() As CIndexedCoeffs, ObjectiveCoeffsBase() As Double
          ExtractModelInfo ValueBase, ObjectiveCoeffsBase
          
          Dim ValueOne() As CIndexedCoeffs, ObjectiveCoeffsOne() As Double
2062      s.BuildModelFromSolverData LinearityOffset:=1, ShouldCheckLinearity:=False
          ExtractModelInfo ValueOne, ObjectiveCoeffsOne
              
          Dim ValueTen() As CIndexedCoeffs, ObjectiveCoeffsTen() As Double
          s.BuildModelFromSolverData LinearityOffset:=10, ShouldCheckLinearity:=False
          ExtractModelInfo ValueTen, ObjectiveCoeffsTen
          
          Dim RowIsNonLinear() As Boolean
2054      If s.NumRows > 0 Then ReDim RowIsNonLinear(s.NumRows) As Boolean

          'Go through each row and check each coefficient individually. if it is not within the tolerance the its nonlinear
          Dim row As Long
          For row = 1 To s.NumRows
2082          RowIsNonLinear(row) = False
              
              Dim FirstVar As Boolean
2083          FirstVar = True

              Dim NumEntries As Long
2087          NumEntries = Max(ValueBase(row).Count, ValueOne(row).Count, ValueTen(row).Count)
              Dim i As Long
2088          For i = 1 To NumEntries
                  Dim VarIndex As Long
                  VarIndex = 0
                  
2089              If TestExistenceOfEntry(ValueBase(row), ValueOne(row), ValueTen(row), i) Then
                      'do a ratio test
2090                  If RatioTest(ValueBase(row).Coefficient(i), ValueOne(row).Coefficient(i)) Or _
                         RatioTest(ValueBase(row).Coefficient(i), ValueTen(row).Coefficient(i)) Then
2092                      VarIndex = ValueOne(row).Index(i)
2095                  End If
2096              Else
2098                  VarIndex = GetEntry(ValueBase(row), ValueOne(row), ValueTen(row), i)
2102              End If

                  If VarIndex > 0 Then
                      Dim constraint As Long, instance As Long
                      constraint = s.RowToConstraint(row)
                      instance = s.GetConstraintInstance(row, constraint)
                      
                      If FirstVar = True Then
2159                      If s.LHSType(constraint) = SolverInputType.SingleCellRange Then
2160                          NonLinearInformation = NonLinearInformation & vbNewLine & "In the constraint: " & s.ConstraintSummary(constraint) & "," & vbNewLine & "  the model appears to be non-linear in the decision variables: "
2161                      Else
                              NonLinearInformation = NonLinearInformation & vbNewLine & "In instance " & instance & " of the constraint: " & s.ConstraintSummary(constraint) & "," & vbNewLine & "  the model appears to be non-linear in the following decision variables: "
2163                      End If
                          FirstVar = False
2164                  Else
2165                      NonLinearInformation = NonLinearInformation & ", "
2166                  End If
                      NonLinearInformation = NonLinearInformation & s.VarNames(VarIndex)
                      
2101                  RowIsNonLinear(row) = True
                  End If
2103          Next i
2105      Next row

          Dim ObjectiveIsNonLinear As Boolean
2109      ObjectiveIsNonLinear = False
2110      For i = 1 To s.NumVars
2111          If RatioTest(ObjectiveCoeffsBase(i), ObjectiveCoeffsOne(i)) Or _
                 RatioTest(ObjectiveCoeffsBase(i), ObjectiveCoeffsTen(i)) Then
2112              If Not ObjectiveIsNonLinear Then
2113                  ObjectiveIsNonLinear = True
2114                  NonLinearInformation = NonLinearInformation & vbNewLine & vbNewLine & "The objective function is nonlinear in the following variables: "
2115              Else
2116                  NonLinearInformation = NonLinearInformation & " , "
2117              End If
                  NonLinearInformation = NonLinearInformation & s.VarNames(i)
2118          End If
2119      Next i

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

Sub DoQuickLinearityCheck()
' Returns false if a full check was performed by the user, meaning the model result is no longer valid.
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim InteractiveStatus As Boolean
          InteractiveStatus = Application.Interactive
          
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
2233          ObservedObjValue = s.ObjRange.Value2
          End If
2236      CalculatedObjValue = CalcObjFnValue(DecisionVariableValues)

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

