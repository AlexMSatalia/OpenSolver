Attribute VB_Name = "SolverModelDiff"
Option Explicit

' Highlight all constraints (and the objective) that are non-linear using our standard model highlighting, but showing only individual cells, not ranges
Sub HighlightNonLinearities(RowIsNonLinear() As Boolean, ObjectiveIsNonLinear As Boolean, s As COpenSolver)
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
2034      For row = 1 To s.NumRows
2035          If RowIsNonLinear(row) Then
                  Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
2036              s.GetConstraintFromRow row, constraint, instance
2037              s.GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
2038              RHSstring = StripWorksheetNameAndDollars(RHSstring, LHSCellRange.Worksheet) ' Strip any worksheet name and $'s from the RHS (useful if it is a formula)
                  Dim RHSisRange As Boolean
2039              RHSisRange = s.RHSType(constraint) = MultiCellRange Or s.RHSType(constraint) = SingleCellRange
2040              HighlightConstraint LHSCellRange.Worksheet, LHSCellRange, RHSCellRange, RHSstring, s.Relation(row), 0  ' Show either a value or a formula
2041          End If
2042      Next row
2043      If ObjectiveIsNonLinear Then
2047          AddObjectiveHighlighting s.ObjRange, s.ObjectiveSense, s.ObjectiveTargetValue
2048      End If

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("SolverModelDiff", "HighlightNonLinearities") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Sub CheckLinearityOfModel(s As COpenSolver)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim row As Long, i As Long, RowIsNonLinear() As Boolean
          Dim ValueZero() As CIndexedCoeffs, ValueOne() As CIndexedCoeffs, ValueTen() As CIndexedCoeffs, OriginalSolutionValues() As Variant
          Dim NonLinearInformation As String
          Dim ObjectiveCoeffsZero() As Double, ObjectiveCoeffsOne() As Double, ObjectiveCoeffsTen() As Double
          Dim ObjectiveFunctionConstantZero As Double, ObjectiveFunctionConstantOne As Double, ObjectiveFunctionConstantTen As Double
          
2049      ReDim SolutionValues(s.numVars) As Double
2050      If s.NumRows > 0 Then ReDim Preserve ValueZero(s.NumRows) As CIndexedCoeffs
          
2051      NonLinearInformation = ""
          
          ' Remember the original decision variable values (in a variant array to handle multiple areas)
2052      OriginalSolutionValues = GetRangeValues(s.AdjustableCells)
          
2053      ReDim ObjectiveCoeffsZero(s.numVars) As Double, ObjectiveCoeffsOne(s.numVars) As Double, ObjectiveCoeffsTen(s.numVars) As Double
2054      If s.NumRows > 0 Then ReDim RowIsNonLinear(s.NumRows) As Boolean
          'Build each matrix where the decision variables start at zero (ValueZero()), one (ValueOne()) and ten (ValueTen())
2055      For row = 1 To s.NumRows
2056          Set ValueZero(row) = s.SparseA(row).Clone
2057      Next row
2058      For i = 1 To s.numVars
2059          ObjectiveCoeffsZero(i) = s.CostCoeffs(i)
2060      Next i
2061      ObjectiveFunctionConstantZero = s.ObjectiveFunctionConstant
          
2062      s.BuildModelFromSolverData 1
2063      If s.NumRows > 0 Then ReDim Preserve ValueOne(s.NumRows) As CIndexedCoeffs
2064      For row = 1 To s.NumRows
2065          Set ValueOne(row) = s.SparseA(row).Clone
2066      Next row
2067      For i = 1 To s.numVars
2068          ObjectiveCoeffsOne(i) = s.CostCoeffs(i)
2069      Next i
2070      ObjectiveFunctionConstantOne = s.ObjectiveFunctionConstant
          
2071      s.BuildModelFromSolverData 10
2072      If s.NumRows > 0 Then ReDim Preserve ValueTen(s.NumRows) As CIndexedCoeffs
2073      For row = 1 To s.NumRows
2074          Set ValueTen(row) = s.SparseA(row).Clone
2075      Next row
2076      For i = 1 To s.numVars
2077          ObjectiveCoeffsTen(i) = s.CostCoeffs(i)
2078      Next i
2079      ObjectiveFunctionConstantTen = s.ObjectiveFunctionConstant
          
          Dim constraint As Long, instance As Long
2080      constraint = 1
          
          'TODO: These tests should not have an AND, and is the model build code valid if we just shift the zero point?
          
          Dim NumEntries As Long, ValueZeroCounter As Long, ValueOneCounter As Long, ValueTenCounter As Long
          Dim FirstVar As Boolean, NonLinearityCount As Long
          'Go through each row and check each coefficient individually. if it is not within the tolerance the its nonlinear
2081      For row = 1 To s.NumRows
2082          RowIsNonLinear(row) = False
              
              'This is used to display the constriant
2083          FirstVar = True
2084          ValueZeroCounter = ValueZero(row).Count
2085          ValueOneCounter = ValueOne(row).Count
2086          ValueTenCounter = ValueTen(row).Count
              'find out how many variables it dependent on
2087          NumEntries = Max(ValueZeroCounter, ValueOneCounter, ValueTenCounter)
2088          For i = 1 To NumEntries
2089              If TestExistanceOfEntry(ValueZero(row), ValueOne(row), ValueTen(row), i) Then
                      'do a ratio test
2090                  If Abs(ValueOne(row).Coefficient(i) - ValueZero(row).Coefficient(i)) / (1 + Abs(ValueOne(row).Coefficient(i))) > EPSILON _
                      And Abs(ValueOne(row).Coefficient(i) - ValueTen(row).Coefficient(i)) / (1 + Abs(ValueOne(row).Coefficient(i))) > EPSILON Then
2091                      s.GetConstraintFromRow row, constraint, instance
2092                      If NonLinearityCount <= 10 Then AddNonLinearInfoToString s, ValueOne(row).Index(i), NonLinearInformation, FirstVar, constraint, instance
2093                      FirstVar = False
2094                      RowIsNonLinear(row) = True
2095                  End If
2096              Else
2097                  s.GetConstraintFromRow row, constraint, instance
                      Dim VariableIndex As Long
2098                  VariableIndex = GetEntry(ValueZero(row), ValueOne(row), ValueTen(row), i)
2099                  If NonLinearityCount <= 10 Then AddNonLinearInfoToString s, VariableIndex, NonLinearInformation, FirstVar, constraint, instance
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
2110      For i = 1 To s.numVars
2111          If Abs(ObjectiveCoeffsZero(i) - ObjectiveCoeffsOne(i)) / (1 + Abs(ObjectiveCoeffsZero(i))) > EPSILON _
              And Abs(ObjectiveCoeffsOne(i) - ObjectiveCoeffsTen(i)) / (1 + Abs(ObjectiveCoeffsOne(i))) > EPSILON Then
2112              If Not ObjectiveIsNonLinear Then
2113                  ObjectiveIsNonLinear = True
2114                  NonLinearInformation = NonLinearInformation & vbNewLine & vbNewLine & "The objective function is nonlinear in the following variables: " & s.VarNames(i)
2115              Else
2116                  NonLinearInformation = NonLinearInformation & " , " & s.VarNames(i)
2117              End If
                   
2118          End If
2119      Next i
          
          'Put the solution back on the sheet
2120      SetRangeValues s.AdjustableCells, OriginalSolutionValues

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
2141          HighlightNonLinearities RowIsNonLinear, ObjectiveIsNonLinear, s
2142      End If

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("SolverModelDiff", "CheckLinearityOfModel") Then Resume
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

          If Not ReportError("SolverModelDiff", "TestExistanceOfEntry") Then Resume
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

Sub AddNonLinearInfoToString(s As COpenSolver, var As Long, NonLinearInformation As String, FirstVar As Boolean, constraint As Long, instance As Long)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

2158      If FirstVar = True Then
2159          If s.LHSType(constraint) = SolverInputType.SingleCellRange Then
2160              NonLinearInformation = NonLinearInformation & vbNewLine & "In the constraint: " & s.ConstraintSummary(constraint) & "," & vbNewLine & "  the model appears to be non-linear in the decision variables: " & s.VarNames(var)
2161          Else
                  NonLinearInformation = NonLinearInformation & vbNewLine & "In instance " & instance & " of the constraint: " & s.ConstraintSummary(constraint) & "," & vbNewLine & "  the model appears to be non-linear in the following decision variables: " & s.VarNames(var)
2163          End If
2164      Else
2165          NonLinearInformation = NonLinearInformation & ", " & s.VarNames(var)
2166      End If

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("SolverModelDiff", "AddNonLinearInfoToString") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Sub QuickLinearityCheck(fullLinearityCheckWasPerformed As Boolean, s As COpenSolver)
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
2180      If s.NumRows > 0 Then ReDim RowIsNonLinear(s.NumRows) As Boolean
          
2181      If Not ForceCalculate("Warning: The worksheet calculation did not complete during the linearity test, and so the test may not be correct. Would you like to retry?") Then GoTo ExitSub
          
          ' Get all the decision variable values off the sheet
          Dim DecisionVariableValues() As Double
2184      ReDim DecisionVariableValues(s.numVars)
2185      DecisionVariableValues = s.GetDecisionVariableValuesOffSheet
          
2186      NonLinearInfo = ""
          Dim NonLinearityCount As Long
          
2187      row = 1
2188      For constraint = 1 To s.NumConstraints
2189          If Not s.LHSRange(constraint) Is Nothing Then ' Skip INT and BINARY constraint
                  ' Get current value(s) for LHS and RHS of this constraint off the sheet
2190              s.GetCurrentConstraintValues constraint, CurrentLHSValues, CurrentRHSValues
                  
2191              If s.RHSType(constraint) <> SolverInputType.MultiCellRange Then
2192                  SolutionValueRHS = CurrentRHSValues
2193              End If
                  
                  Dim instance As Long
2194              instance = 0
2195              For i = 1 To UBound(CurrentLHSValues, 1)
2196                  For j = 1 To UBound(CurrentLHSValues, 2)
2197                      instance = instance + 1
2198                      SolutionValueLHS = CurrentLHSValues(i, j)
2199                      If s.RHSType(constraint) = SolverInputType.MultiCellRange Then
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
2208                      ExpectedValue = s.SparseA(row).Evaluate_RecordPrecision(DecisionVariableValues, maxValueInCalculation) - s.RHS(row)
2209                      If Abs(s.RHS(row)) > maxValueInCalculation Then maxValueInCalculation = Abs(s.RHS(row))
          
                          ' do a ratio test
2210                      If Abs(ExpectedValue - SolutionValue) / (1 + Abs(ExpectedValue)) > Max(EPSILON, EPSILON * maxValueInCalculation) Then
2211                          If NonLinearInfo = "" Then NonLinearInfo = "The following constraint(s) do not appear to be linear: "
2212                          If NonLinearityCount <= 10 Then
2213                              NonLinearInfo = NonLinearInfo & vbNewLine & s.ConstraintSummary(constraint)
                              
                                  Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
2214                              s.GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
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
          If s.ObjRange Is Nothing Then
              ObservedObjValue = 0
          Else
2233          ObservedObjValue = s.ObjRange.Value2
          End If
          
2236      CalculatedObjValue = s.CalcObjFnValue(DecisionVariableValues)
2237      ObjectiveIsNonLinear = Abs(CalculatedObjValue - ObservedObjValue) / (1 + Abs(CalculatedObjValue)) > EPSILON
2238      If ObjectiveIsNonLinear Then
2239         NonLinearInfo = "The objective function is not linear." & vbNewLine & vbNewLine & NonLinearInfo
2240      End If
          
          'Set the userform up and display any information on nonlinear constraints
2241      If NonLinearInfo <> "" Then
2242          s.SolveStatus = NotLinear
2243          If Not s.MinimiseUserInteraction Then
                  NonLinearInfo = "WARNING : " & vbNewLine & TrimBlankLines(NonLinearInfo)
                  frmNonlinear.SetLinearityResult NonLinearInfo, True
2263              frmNonlinear.Show
              
                  'showing the nonlinear constraints
2264              If frmNonlinear.chkHighlight.value = True Then
2265                  HighlightNonLinearities RowIsNonLinear, ObjectiveIsNonLinear, s
2266              End If
2267              If frmNonlinear.chkFullCheck.value = True Then
                      'Full linearity check run
2268                  CheckLinearityOfModel s
2269                  fullLinearityCheckWasPerformed = True
2270              End If
2271          End If
2272      End If

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("SolverModelDiff", "QuickLinearityCheck") Then Resume
          RaiseError = True
          GoTo ExitSub
          
End Sub
Sub WriteConstraintListToSheet(r As Range, s As COpenSolver)
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
1921      For row = 1 To s.NumRows
1922          s.GetConstraintFromRow row, constraint, instance  ' Which Excel constraint are we in, and which instance?
              
              Dim UnusedConstraint As Boolean
1923          UnusedConstraint = s.SparseA(row).Count = 0
              
              Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
1924          s.GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
1925          If Not RHSCellRange Is Nothing Then
1926              RHSstring = RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
1927          Else
1928              RHSstring = ConvertToCurrentLocale(StripWorksheetNameAndDollars(RHSstring, ActiveSheet))
1930          End If
              
              Dim summary As String
1931          summary = IIf(UnusedConstraint, "", "") & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & _
                        RelationEnumToString(s.Relation(row)) & RHSstring & IIf(UnusedConstraint, "", "")
          
1932          r.Cells(row + 1, 1).value = summary
1933          r.Cells(row + 1, 2).Value2 = ZeroIfSmall(s.ShadowPrice(row))
1934          r.Cells(row + 1, 3).Value2 = ZeroIfSmall(s.IncreaseCon(row))
1935          r.Cells(row + 1, 4).Value2 = ZeroIfSmall(s.DecreaseCon(row))
1936      Next row

          'Write the variable duals
1937      row = row + 2
1938      r.Cells(row, 1).Value2 = "Vars"
1939      r.Cells(row, 2).Value2 = "RC"
1940      r.Cells(row, 3).Value2 = "Inc"
1941      r.Cells(row, 4).Value2 = "Dec"
1942      row = row + 1
          Dim i As Long
1943      For i = 1 To s.numVars
1944          r.Cells(row, 1).Value2 = s.VarCell(i)
1945          r.Cells(row, 2).Value2 = ZeroIfSmall(s.ReducedCosts(i))
1946          r.Cells(row, 3).Value2 = ZeroIfSmall(s.IncreaseVar(i))
1947          r.Cells(row, 4).Value2 = ZeroIfSmall(s.DecreaseVar(i))
1948          row = row + 1
1949      Next i

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("SolverModelDiff", "WriteConstraintListToSheet") Then Resume
          RaiseError = True
          GoTo ExitSub

End Sub
Sub WriteConstraintSensitivityTable(nameSheet As String, s As COpenSolver)
'Writes out the sensitivity table on a new page (like the solver sensitivity report)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim Column As Long, row As Long, i As Long
          
2273      Sheets(nameSheet).Cells(1, 1) = "OpenSolver Sensitivity Report - " & s.Solver.ShortName
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
          
          'put the values into the variable table
2286      For i = 1 To s.numVars
2287          Sheets(nameSheet).Cells(row, Column) = s.VarCell(i)
2288          Sheets(nameSheet).Cells(row, Column + 2) = ZeroIfSmall(s.FinalVarValue(i))
2289          Sheets(nameSheet).Cells(row, Column + 3) = ZeroIfSmall(s.ReducedCosts(i))
2290          Sheets(nameSheet).Cells(row, Column + 4) = ZeroIfSmall(s.CostCoeffs(i))
2291          Sheets(nameSheet).Cells(row, Column + 5) = ZeroIfSmall(s.IncreaseVar(i))
2292          Sheets(nameSheet).Cells(row, Column + 6) = ZeroIfSmall(s.DecreaseVar(i))
2293          Sheets(nameSheet).Cells(row, Column + 1) = findName(sheet, s.VarCell(i))
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
2305      For i = 1 To s.NumRows
2306          Sheets(nameSheet).Cells(row, Column + 2) = ZeroIfSmall(s.FinalValue(i))
2307          Sheets(nameSheet).Cells(row, Column + 3) = ZeroIfSmall(s.ShadowPrice(i))
2308          Sheets(nameSheet).Cells(row, Column + 4) = ZeroIfSmall(s.RHS(i))
2309          Sheets(nameSheet).Cells(row, Column + 5) = ZeroIfSmall(s.IncreaseCon(i))
2310          Sheets(nameSheet).Cells(row, Column + 6) = ZeroIfSmall(s.DecreaseCon(i))
              'This finds the range for cells of each constraint (similar to WriteConstraintListToSheet)
2311          s.GetConstraintFromRow i, constraint, instance  ' Which Excel constraint are we in, and which instance?
              Dim UnusedConstraint As Boolean
2312          UnusedConstraint = s.SparseA(i).Count = 0
              Dim LHSCellRange As Range, RHSCellRange As Range, RHSstring As String
2313          s.GetConstraintInstanceData constraint, instance, LHSCellRange, RHSCellRange, RHSstring
2314          If Not RHSCellRange Is Nothing Then
2315              RHSstring = RHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
2316          Else
2317              RHSstring = ConvertToCurrentLocale(StripWorksheetNameAndDollars(RHSstring, ActiveSheet))
2319          End If
              Dim summary As String
2320          summary = IIf(UnusedConstraint, "", "") & LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & _
                        RelationEnumToString(s.Relation(i)) & RHSstring & IIf(UnusedConstraint, "", "")
              'Cell Range for each constraint
2321          Sheets(nameSheet).Cells(row, Column).value = summary
              'Finds the nearest name for the constraint
2322          Sheets(nameSheet).Cells(row, Column + 1) = findName(sheet, LHSCellRange.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False))
2323          row = row + 1
2324      Next i

          'Format the sensitivity table
2325      FormatSensitivityTable nameSheet, row, s.numVars

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("SolverModelDiff", "WriteConstraintSensitivityTable") Then Resume
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
          ActiveWindow.DisplayGridlines = False
          
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
          If Not ReportError("SolverModelDiff", "FormatSensitivityTable") Then Resume
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
2429              CellValue = CStr(Sheets(sheet).Cells(row, i).value)
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
2443      LHSName = CStr(Cells(row, i + 1).value)
2444      AboveName = CStr(Cells(j + 1, col).value)
          
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
          If Not ReportError("SolverModelDiff", "findName") Then Resume
          RaiseError = True
          GoTo ExitFunction

End Function

