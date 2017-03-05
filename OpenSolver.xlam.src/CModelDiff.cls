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
1               RaiseError = False
2               On Error GoTo ErrorHandler
                
3               Set s = OpenSolver
                
4               If Not BuildConstantTerms() Then GoTo ExitFunction
5               If Not BuildVariableTerms(LinearityOffset) Then GoTo ExitFunction
                
6               If Not QuickFeasibilityCheck() Then GoTo ExitFunction
                
7               If ShouldCheckLinearity Then
8                   DoQuickLinearityCheck
9               End If
                
10              s.AdjustableCells.Value2 = 0
11              ProcessSolverModel = True

ExitFunction:
12              Application.StatusBar = False
13              If RaiseError Then RethrowError
14              Exit Function
                
ErrorHandler:
15              If Not ReportError("CModelDiff", "ProcessSolverModel") Then Resume
16              RaiseError = True
17              GoTo ExitFunction
End Function

Private Function BuildConstantTerms() As Boolean
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          'Create the rhs vector by resetting cells to zero
3         s.AdjustableCells.Value2 = 0
          ' Recalculate the spreadsheet
4         If Not ForceCalculate("Warning: The worksheet calculation did not complete, and so the model may not be generated correctly. Would you like to retry?") Then GoTo ExitFunction

          ' TODO Handle errors in the original spreadsheet values
          Dim ZeroedLHSValues As Variant, ZeroedRHSValues As Variant
          Dim row As Long, constraint As Long, i As Long, j As Long
5         row = 1
6         For constraint = 1 To s.NumConstraints
7             If Not s.LHSRange(constraint) Is Nothing Then
                  ' Get current value(s) for LHS and RHS of this constraint off the sheet. LHS is always an array (even if 1x1)
8                 s.GetCurrentConstraintValues constraint, ZeroedLHSValues, ZeroedRHSValues
9                 If s.RHSType(constraint) = SolverInputType.MultiCellRange Then 'does the constraint refer to more than one cell
10                    For i = 1 To UBound(ZeroedLHSValues, 1)
11                        For j = 1 To UBound(ZeroedLHSValues, 2)
12                            If UBound(ZeroedLHSValues, 1) = UBound(ZeroedRHSValues, 1) Then
13                                s.RHS(row) = ZeroedRHSValues(i, j) - ZeroedLHSValues(i, j)
14                            Else
15                                s.RHS(row) = ZeroedRHSValues(j, i) - ZeroedLHSValues(i, j)
16                            End If
17                            row = row + 1
18                        Next j
19                    Next i
20                Else
21                    For i = 1 To UBound(ZeroedLHSValues, 1)
22                        For j = 1 To UBound(ZeroedLHSValues, 2)
23                            s.RHS(row) = ZeroedRHSValues - ZeroedLHSValues(i, j)
24                            row = row + 1
25                        Next j
26                    Next i
27                End If
28            End If
29        Next constraint

30        If Not s.ObjRange Is Nothing Then
31            s.ObjectiveFunctionConstant = s.GetCurrentObjectiveValue()
              ' Add constants to Target
32            If s.ObjectiveSense = TargetObjective Then
                  ' Target value needs to be incremented by any constants in the objective
33                s.ObjectiveTargetValue = s.ObjectiveTargetValue - s.ObjectiveFunctionConstant
34            End If
35        End If
          
36        BuildConstantTerms = True

ExitFunction:
37        Application.StatusBar = False
38        If RaiseError Then RethrowError
39        Exit Function
          
ErrorHandler:
40        If Not ReportError("CModelDiff", "BuildConstantTerms") Then Resume
41        RaiseError = True
42        GoTo ExitFunction
End Function

Private Function BuildVariableTerms(LinearityOffset As Double, Optional BinaryOffset As Double = 0) As Boolean
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          ' Initialize
          Dim row As Long
3         For row = 1 To s.NumRows
4             Set s.SparseA(row) = New CIndexedCoeffs
5         Next row
6         Set s.CostCoeffs = New CIndexedCoeffs

7         s.AdjustableCells.Value2 = LinearityOffset
8         If Not s.BinaryCellsRange Is Nothing Then s.BinaryCellsRange.Value2 = BinaryOffset
          
9         If Not ForceCalculate("Warning: The worksheet calculation did not complete, and so the model may not be generated correctly. Would you like to retry?") Then GoTo ExitFunction

          Dim LHSOriginalValues() As Variant, RHSOriginalValues() As Variant
10        ReDim LHSOriginalValues(s.NumConstraints)
11        ReDim RHSOriginalValues(s.NumConstraints)
          Dim constraint As Long
12        For constraint = 1 To s.NumConstraints
13            If Not s.LHSRange(constraint) Is Nothing Then
14                s.GetCurrentConstraintValues constraint, LHSOriginalValues(constraint), RHSOriginalValues(constraint)
15            End If
16        Next constraint
          
          Dim ObjPresent As Boolean, OriginalObjectiveValue As Double
17        ObjPresent = Not s.ObjRange Is Nothing
18        If ObjPresent Then OriginalObjectiveValue = s.GetCurrentObjectiveValue()
          
          ' Increment each decision variable by one to see what constraints are dependent on it and by what amount
          Dim var As Long, AdjCell As Range, coeff As Double
19        var = 0
20        For Each AdjCell In s.AdjustableCells
              ' Count how many variables we've processed, let the user know that the program hasn't crashed!
21            var = var + 1
22            UpdateStatusBar "OpenSolver: Setting Up Problem... " & var & "/" & s.NumVars & " vars, " & s.NumRows & " rows."
              
              ' We start binaries at 0 to test 0 and 1, and all others at the specified offset value
              Dim BaseValue As Double, NewValue As Double
23            If ProperIntersect(AdjCell, s.BinaryCellsRange) Is Nothing Then
24                BaseValue = LinearityOffset
25                NewValue = LinearityOffset + 1
26            Else
27                BaseValue = BinaryOffset
28                NewValue = 1 - BinaryOffset
29            End If
30            AdjCell.Value2 = NewValue

31            If Not ForceCalculate("Warning: The worksheet calculation did not complete, and so the model may not be generated correctly. Would you like to retry?") Then GoTo ExitFunction

              ' Easy part first - the objective function value change
32            If ObjPresent Then
33                coeff = s.GetCurrentObjectiveValue - OriginalObjectiveValue
34                If Abs(coeff) > EPSILON Then
35                    s.CostCoeffs.Add var, coeff
36                End If
37            End If

              ' Now the hard (slow) bit - the constraint changes
38            row = 1
39            For constraint = 1 To s.NumConstraints

                  ' Check to see what is different and add rows to SparseA
40                If Not s.LHSRange(constraint) Is Nothing Then ' skip Binary and Integer constraints

                      Dim OriginalLHSValues As Variant
                      Dim OriginalRHSValues As Variant
41                    OriginalLHSValues = LHSOriginalValues(constraint)
42                    OriginalRHSValues = RHSOriginalValues(constraint)

                      ' Get current value(s) for LHS and RHS of this constraint off the sheet. LHS is always an array (even if 1x1)
                      Dim CurrentLHSValues As Variant
                      Dim CurrentRHSValues As Variant
43                    s.GetCurrentConstraintValues constraint, CurrentLHSValues, CurrentRHSValues

                      Dim i As Long, j As Long
44                    For i = 1 To UBound(OriginalLHSValues, 1) ' number of rows
45                        For j = 1 To UBound(OriginalLHSValues, 2) ' number of columns
46                            Select Case s.RHSType(constraint)
                              Case SolverInputType.MultiCellRange
                                  '-----------------------------------------------------------
                                  'Making it work for column LHS with row RHS and vice versa
                                  '-----------------------------------------------------------
47                                If UBound(OriginalLHSValues, 1) = UBound(OriginalRHSValues, 1) Then
48                                    coeff = (CurrentLHSValues(i, j) - OriginalLHSValues(i, j)) - (CurrentRHSValues(i, j) - OriginalRHSValues(i, j))
49                                Else
50                                    coeff = (CurrentLHSValues(i, j) - OriginalLHSValues(i, j)) - (CurrentRHSValues(j, i) - OriginalRHSValues(j, i))
51                                End If
52                            Case SolverInputType.Formula, SolverInputType.SingleCellRange
53                                coeff = (CurrentLHSValues(i, j) - OriginalLHSValues(i, j)) - (CurrentRHSValues - OriginalRHSValues)
54                            Case SolverInputType.constant
55                                coeff = (CurrentLHSValues(i, j) - OriginalLHSValues(i, j))
56                            End Select
                              ' Ignore zero (or near-zero) coefficients
57                            If Abs(coeff) > EPSILON Then
58                                s.SparseA(row).Add var, coeff
59                            End If
60                            row = row + 1
61                        Next j
62                    Next i
63                End If
64            Next constraint

65            AdjCell.Value2 = BaseValue
66        Next AdjCell
          
67        BuildVariableTerms = True
          
ExitFunction:
68        Application.StatusBar = False
69        If RaiseError Then RethrowError
70        Exit Function

ErrorHandler:
71        If Not ReportError("COpenSolver", "BuildVariableTerms") Then Resume
72        RaiseError = True
73        GoTo ExitFunction
End Function

Function QuickFeasibilityCheck() As Boolean
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          ' Check if any empty constraints are feasible
          Dim row As Long
3         For row = 1 To s.NumRows
4             If s.SparseA(row).Count = 0 Then
                  ' We have a constraint that does not vary with the decision variables; check it is satisfied
                  Dim constraint As Long, instance As Long
5                 constraint = s.RowToConstraint(row)
6                 instance = s.GetConstraintInstance(row, constraint)
                  
                  Dim rel As RelationConsts, value As Double
7                 rel = s.Relation(constraint)
8                 value = s.RHS(row)
                  
9                 If (rel = RelationEQ And Abs(value) > EPSILON) Or _
                     (rel = RelationGE And value > EPSILON) Or _
                     (rel = RelationLE And value < -EPSILON) Then
                      
                      Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
10                    s.GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
                      ' If the RHS is a range, we show its address; if not, RHSString contains the RHS's constant or formula
11                    If Not RHSCellRange Is Nothing Then RHSstring = RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)

                      Dim LHSValue As Double, RHSValue As Double
12                    LHSValue = LHSCellRange.Value2
13                    RHSValue = s.RHS(row) + LHSValue
                      
14                    s.SolveStatus = OpenSolverResult.Infeasible
15                    s.SolveStatusString = "Infeasible"
16                    s.SolveStatusComment = "The model contains a constraint '" & s.ConstraintSummary(constraint) & "' for which instance " & _
                                             instance & " does not depend on the decision variables and is not satisfied." & vbCrLf & _
                                             "Constraint specifies: " & _
                                             "LHS=" & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "=" & LHSValue & _
                                             " " & RelationEnumToString(rel) & " " & _
                                             "RHS=" & StripWorksheetNameAndDollars(RHSstring, s.sheet) & "=" & RHSValue
17                    GoTo ExitFunction
18                End If
19            End If
20        Next row
          
          ' Check objective if it is a target and all zeros
21        If s.ObjectiveSense = TargetObjective Then
22            If s.CostCoeffs.Count = 0 And _
                 Abs(s.ObjectiveTargetValue - s.ObjectiveFunctionConstant) > EPSILON Then
23                s.SolveStatus = OpenSolverResult.Infeasible
24                s.SolveStatusString = "Infeasible"
25                s.SolveStatusComment = "The objective cell does not depend on the adjustable cells " & _
                                         "and has a value different to the specified target value. " & _
                                         "This means the target value cannot be attained."
26                GoTo ExitFunction
27            End If
28        End If
          
29        QuickFeasibilityCheck = True

ExitFunction:
30        If RaiseError Then RethrowError
31        Exit Function

ErrorHandler:
32        If Not ReportError("CModelDiff", "QuickFeasibilityCheck") Then Resume
33        RaiseError = True
34        GoTo ExitFunction
End Function

' =========================================
' Linearity check functions
' =========================================

' Highlight all constraints (and the objective) that are non-linear using our standard model highlighting, but showing only individual cells, not ranges
Sub HighlightNonLinearities(RowIsNonLinear() As Boolean, ObjectiveIsNonLinear As Boolean)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If SheetHasOpenSolverHighlighting(s.sheet) Then HideSolverModel s.sheet
4         DeleteOpenSolverShapes s.sheet

5         InitialiseHighlighting

          Dim constraint As Long, row As Long, instance As Long
6         For row = 1 To s.NumRows
7             If RowIsNonLinear(row) Then
8                 constraint = s.RowToConstraint(row)
9                 instance = s.GetConstraintInstance(row, constraint)

                  Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
10                s.GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring

11                RHSstring = StripWorksheetNameAndDollars(RHSstring, LHSCellRange.Worksheet) ' Strip any worksheet name and $'s from the RHS (useful if it is a formula)
12                HighlightConstraint LHSCellRange.Worksheet, LHSCellRange, RHSCellRange, RHSstring, s.Relation(constraint), 0  ' Show either a value or a formula
13            End If
14        Next row

15        If ObjectiveIsNonLinear Then
16            AddObjectiveHighlighting s.ObjRange, s.ObjectiveSense, s.ObjectiveTargetValue
17        End If

ExitSub:
18        If RaiseError Then RethrowError
19        Exit Sub

ErrorHandler:
20        If Not ReportError("CModelDiff", "HighlightNonLinearities") Then Resume
21        RaiseError = True
22        GoTo ExitSub
End Sub

Private Sub ExtractModelInfo(ByRef value() As CIndexedCoeffs, ByRef ObjectiveCoeffs As CIndexedCoeffs)
1         If s.NumRows > 0 Then ReDim Preserve value(s.NumRows) As CIndexedCoeffs

          Dim row As Long
2         For row = 1 To s.NumRows
3             Set value(row) = s.SparseA(row).Clone
4         Next row

5         Set ObjectiveCoeffs = s.CostCoeffs.Clone
End Sub

Private Function RatioTest(Value1 As Double, Value2 As Double, Optional Threshold As Double = EPSILON) As Boolean
1               RatioTest = Abs(Value1 - Value2) / (1 + Abs(Value1)) > Threshold
End Function

Private Sub DoFullLinearityCheck()
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim InteractiveStatus As Boolean
3         InteractiveStatus = Application.Interactive
          
          Dim NonLinearInformation As String
4         NonLinearInformation = vbNullString
          
          'Build each matrix where the decision variables start at the base linearity offset (ValueBase()), one (ValueOne()) and ten (ValueTen())
          Dim ValueBase() As CIndexedCoeffs, ObjectiveCoeffsBase As CIndexedCoeffs
5         ExtractModelInfo ValueBase, ObjectiveCoeffsBase
          
          Dim ValueOne() As CIndexedCoeffs, ObjectiveCoeffsOne As CIndexedCoeffs
6         BuildVariableTerms LinearityOffset:=1
7         ExtractModelInfo ValueOne, ObjectiveCoeffsOne
              
          Dim ValueTen() As CIndexedCoeffs, ObjectiveCoeffsTen As CIndexedCoeffs
8         BuildVariableTerms LinearityOffset:=10, BinaryOffset:=1
9         ExtractModelInfo ValueTen, ObjectiveCoeffsTen
          
          Dim RowIsNonLinear() As Boolean
10        If s.NumRows > 0 Then ReDim RowIsNonLinear(s.NumRows) As Boolean

          Dim row As Long
11        For row = 1 To s.NumRows
12            RowIsNonLinear(row) = False
              
              Dim NonLinearConVars As String
13            NonLinearConVars = CompareSparseVectors(ValueBase(row), ValueOne(row), ValueTen(row))
14            If Len(NonLinearConVars) > 0 Then
                  Dim constraint As Long, instance As Long
15                constraint = s.RowToConstraint(row)
16                instance = s.GetConstraintInstance(row, constraint)
                  
17                If s.LHSType(constraint) = SolverInputType.SingleCellRange Then
18                    NonLinearInformation = NonLinearInformation & vbNewLine & "In the constraint: " & s.ConstraintSummary(constraint) & "," & vbNewLine & "  the model appears to be non-linear in the decision variables: "
19                Else
20                    NonLinearInformation = NonLinearInformation & vbNewLine & "In instance " & instance & " of the constraint: " & s.ConstraintSummary(constraint) & "," & vbNewLine & "  the model appears to be non-linear in the following decision variables: "
21                End If
22                NonLinearInformation = NonLinearInformation & NonLinearConVars
23                RowIsNonLinear(row) = True
24            End If
25        Next row

          Dim ObjectiveIsNonLinear As Boolean, NonLinearObjVars As String
26        ObjectiveIsNonLinear = False
27        NonLinearObjVars = CompareSparseVectors(ObjectiveCoeffsBase, ObjectiveCoeffsOne, ObjectiveCoeffsTen)
28        If Len(NonLinearObjVars) > 0 Then
29            ObjectiveIsNonLinear = True
30            NonLinearInformation = NonLinearInformation & vbNewLine & vbNewLine & _
                                     "The objective function is nonlinear in the following variables: " & _
                                     NonLinearObjVars
31        End If

32        NonLinearInformation = TrimBlankLines(NonLinearInformation)
33        If Len(NonLinearInformation) = 0 Then
34            NonLinearInformation = "There have been no instances of nonlinearity found in this model. Some models can generate warnings of non-linearity " & _
                                     "because of numerical errors that accumulate in the spreadsheet. OpenSolver's non-linearity check can be disabled under OpenSolver's " & _
                                     "Options settings."
35        End If
          
          'display dialog to user
          Dim frmNonlinear As FNonlinear
36        Set frmNonlinear = New FNonlinear
37        frmNonlinear.SetLinearityResult NonLinearInformation, False
          
38        Application.Interactive = True
39        frmNonlinear.Show
40        Application.Interactive = InteractiveStatus
          
41        If frmNonlinear.chkHighlight.value = True Then
42            HighlightNonLinearities RowIsNonLinear, ObjectiveIsNonLinear
43        End If
44        Unload frmNonlinear

ExitSub:
45        Application.Interactive = InteractiveStatus
46        If RaiseError Then RethrowError
47        Exit Sub

ErrorHandler:
48        If Not ReportError("CModelDiff", "DoFullLinearityCheck") Then Resume
49        RaiseError = True
50        GoTo ExitSub
End Sub

Function TestExistenceOfEntry(ValueBase As CIndexedCoeffs, ValueOne As CIndexedCoeffs, ValueTen As CIndexedCoeffs, i As Long) As Boolean
          'Check if this index exists in all three models
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If ValueBase.Index(i) <> 0 And ValueOne.Index(i) <> 0 And ValueTen.Index(i) <> 0 Then
4             TestExistenceOfEntry = True
5         End If

ExitFunction:
6         If RaiseError Then RethrowError
7         Exit Function

ErrorHandler:
8         If Err.Number = 9 Then
9             TestExistenceOfEntry = False
10            Resume ExitFunction
11        End If

12        If Not ReportError("CModelDiff", "TestExistenceOfEntry") Then Resume
13        RaiseError = True
14        GoTo ExitFunction
End Function

'Return the ith entry from any one of these arrays; it may not exist in all of them
Function GetEntry(ValueBase As CIndexedCoeffs, ValueOne As CIndexedCoeffs, ValueTen As CIndexedCoeffs, i As Long) As Long
1         On Error Resume Next
2         If i <= ValueBase.Count Then
3             GetEntry = ValueBase.Index(i)
4         ElseIf i <= ValueOne.Count Then
5             GetEntry = ValueOne.Index(i)
6         ElseIf i <= ValueTen.Count Then
7             GetEntry = ValueTen.Index(i)
8         End If
End Function

Function CompareSparseVectors(v1 As CIndexedCoeffs, v2 As CIndexedCoeffs, v3 As CIndexedCoeffs) As String
      ' Returns a collection of all variable indices with different coefficients in the vectors
                Dim NumEntries As Long
1               NumEntries = Max(v1.Count, v2.Count, v3.Count)
                
                Dim VarString As String
                
                Dim i As Long
2               For i = 1 To NumEntries
                    Dim VarIndex As Long
3                   VarIndex = 0
                        
4                   If TestExistenceOfEntry(v1, v2, v3, i) Then
                        'do a ratio test
5                       If RatioTest(v1.Coefficient(i), v2.Coefficient(i)) Or _
                           RatioTest(v1.Coefficient(i), v3.Coefficient(i)) Then
6                           VarIndex = v1.Index(i)
7                       End If
8                   Else
9                       VarIndex = GetEntry(v1, v2, v3, i)
10                  End If
                    
11                  If VarIndex > 0 Then
12                      VarString = VarString & ", " & s.VarName(VarIndex)
13                  End If
14              Next i
                
15              CompareSparseVectors = Mid(VarString, 3)  ' Trim leading comma
End Function

Sub DoQuickLinearityCheck()
      ' Returns false if a full check was performed by the user, meaning the model result is no longer valid.
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim InteractiveStatus As Boolean
3         InteractiveStatus = Application.Interactive
          
          ' Do a better check for binary non-linearity - build from 0, check at 1
4         If Not s.BinaryCellsRange Is Nothing Then
5             s.BinaryCellsRange.Value2 = 1
6         End If
          
7         If Not ForceCalculate("Warning: The worksheet calculation did not complete during the linearity test, and so the test may not be correct. Would you like to retry?") Then GoTo ExitSub
          
          ' Get all the decision variable values off the sheet
          Dim DecisionVariableValues() As Double, var As Long, c As Range
8         ReDim DecisionVariableValues(s.NumVars)
9         var = 1
10        For Each c In s.AdjustableCells
11            DecisionVariableValues(var) = c.Value2
12            var = var + 1
13        Next c
          
          Dim NonLinearInfo As String
14        NonLinearInfo = vbNullString

          Dim RowIsNonLinear() As Boolean
15        If s.NumRows > 0 Then ReDim RowIsNonLinear(s.NumRows) As Boolean
          
          Dim constraint As Long, row As Long
16        row = 1
17        For constraint = 1 To s.NumConstraints
18            If Not s.LHSRange(constraint) Is Nothing Then
                  ' Get current value(s) for LHS and RHS of this constraint off the sheet
                  Dim CurrentLHSValues As Variant, CurrentRHSValues As Variant
19                s.GetCurrentConstraintValues constraint, CurrentLHSValues, CurrentRHSValues
                  
                  Dim instance As Long, i As Long, j As Long
20                instance = 0
21                For i = 1 To UBound(CurrentLHSValues, 1)
22                    For j = 1 To UBound(CurrentLHSValues, 2)
23                        instance = instance + 1
                          
                          Dim SolutionValueLHS As Double, SolutionValueRHS As Double, SolutionValue As Double
24                        SolutionValueLHS = CurrentLHSValues(i, j)
25                        If s.RHSType(constraint) = SolverInputType.MultiCellRange Then
                              ' Check whether the LHS and RHS are parallel or perpendicular
26                            If UBound(CurrentLHSValues, 1) = UBound(CurrentRHSValues, 1) Then
27                                SolutionValueRHS = CurrentRHSValues(i, j)
28                            Else
29                                SolutionValueRHS = CurrentRHSValues(j, i)
30                            End If
31                        Else
32                            SolutionValueRHS = CurrentRHSValues
33                        End If
34                        SolutionValue = SolutionValueLHS - SolutionValueRHS
                      
                          ' Find out what we expect the value to be from Ax = b. We track the maximum value we encounter during the calculation
                          ' so that we have some idea of the errors we might expect
                          Dim ExpectedValue As Double, MaxValueInCalculation As Double
35                        ExpectedValue = s.SparseA(row).Evaluate_RecordPrecision(DecisionVariableValues, MaxValueInCalculation) - s.RHS(row)
36                        MaxValueInCalculation = Max(MaxValueInCalculation, Abs(s.RHS(row)))
          
37                        If RatioTest(ExpectedValue, SolutionValue, Threshold:=Max(EPSILON, EPSILON * MaxValueInCalculation)) Then
                              Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
38                            s.GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
                              ' If the RHS is a range, we show its address; if not, RHSString contains the RHS's constant or formula
39                            If Not RHSCellRange Is Nothing Then RHSstring = RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
                              
40                            If Len(NonLinearInfo) = 0 Then NonLinearInfo = "The following constraint(s) do not appear to be linear: "
41                            NonLinearInfo = NonLinearInfo & vbNewLine & s.ConstraintSummary(constraint) & ": instance " & instance & _
                                              ", LHS=" & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & _
                                              ", RHS=" & RHSstring & _
                                              ", " & ExpectedValue & "<>" & SolutionValue
                              
42                            RowIsNonLinear(row) = True
43                        End If
44                        row = row + 1
45                    Next j
46                Next i
47            End If
48        Next constraint

          'Check objective function for linearity
          Dim CalculatedObjValue As Double, ObservedObjValue As Double, ObjectiveIsNonLinear As Boolean
49        If s.ObjRange Is Nothing Then
50            ObservedObjValue = 0
51        Else
52            ObservedObjValue = s.GetCurrentObjectiveValue()
53        End If
54        CalculatedObjValue = s.CostCoeffs.Evaluate_RecordPrecision(DecisionVariableValues, MaxValueInCalculation) + s.ObjectiveFunctionConstant
55        MaxValueInCalculation = Max(MaxValueInCalculation, Abs(s.ObjectiveFunctionConstant))

56        ObjectiveIsNonLinear = RatioTest(CalculatedObjValue, ObservedObjValue)
57        If ObjectiveIsNonLinear Then
58           NonLinearInfo = "The objective function is not linear." & vbNewLine & vbNewLine & NonLinearInfo
59        End If
          
          'Set the userform up and display any information on nonlinear constraints
60        If Len(NonLinearInfo) > 0 Then
61            s.SolveStatus = NotLinear
62            If Not s.MinimiseUserInteraction Then
63                NonLinearInfo = "WARNING : " & vbNewLine & TrimBlankLines(NonLinearInfo)
                  
                  Dim frmNonlinear As FNonlinear
64                Set frmNonlinear = New FNonlinear
65                frmNonlinear.SetLinearityResult NonLinearInfo, True
                  
66                Application.Interactive = True
67                frmNonlinear.Show
68                Application.Interactive = InteractiveStatus
              
                  'Show the nonlinear constraints if requested
69                If frmNonlinear.chkHighlight.value = True Then
70                    HighlightNonLinearities RowIsNonLinear, ObjectiveIsNonLinear
71                End If
                  ' Conduct full linearity check if requested
72                If frmNonlinear.chkFullCheck.value = True Then DoFullLinearityCheck

73                Unload frmNonlinear
74            End If
75        End If

ExitSub:
76        Application.Interactive = InteractiveStatus
77        If RaiseError Then RethrowError
78        Exit Sub

ErrorHandler:
79        If Not ReportError("CModelDiff", "DoQuickLinearityCheck") Then Resume
80        RaiseError = True
81        GoTo ExitSub
          
End Sub

