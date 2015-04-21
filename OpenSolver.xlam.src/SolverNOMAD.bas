Attribute VB_Name = "SolverNOMAD"
Option Explicit

Public OS As COpenSolver
Dim IterationCount As Long

' NOMAD functions
' Do not put error handling in these functions, there is already error-handling in the NOMAD DLL

Sub NOMAD_UpdateVar(X As Variant, Optional BestSolution As Variant = Nothing, Optional Infeasible As Boolean = False)
7115      IterationCount = IterationCount + 1

          ' Update solution
7116      Dim status As String
7117      status = "OpenSolver: Running NOMAD. Iteration " & IterationCount & "."
          ' Check for BestSolution = Nothing
7118      If Not VarType(BestSolution) = 9 Then
              ' Flip solution if maximisation
7119          If OS.ObjectiveSense = MaximiseObjective Then BestSolution = -BestSolution

7120          status = status & " Best solution so far: " & BestSolution
7121          If Infeasible Then
7122              status = status & " (infeasible)"
7123          End If
7124      End If
7125      UpdateStatusBar status

          Dim i As Long, numVars As Long
          'set new variable values on sheet
2452      numVars = UBound(X)
2453      i = 1
          Dim AdjCell As Range
          ' If only one variable is returned, X is treated as a 1D array rather than 2D, so we need to access it
          ' differently.
2454      If numVars = 1 Then
2455          For Each AdjCell In OS.AdjustableCells
2456              AdjCell.Value2 = X(i)
2457              i = i + 1
2458          Next AdjCell
2459      Else
2460          For Each AdjCell In OS.AdjustableCells
2461              AdjCell.Value2 = X(i, 1)
2462              i = i + 1
2463          Next AdjCell
2464      End If
End Sub

Function NOMAD_GetValues() As Variant
          Dim X As Variant, i As Long, j As Long, k As Long, numCons As Variant
2465      numCons = NOMAD_GetNumConstraints()
2466      ReDim X(1 To numCons(0), 1 To 1)
          '====NOMAD only does minimise so need to change objective if it is max====
          ' If no objective, just set a constant.
          ' TODO: fix this to set it based on amount of violation to hunt for feasibility
2467      If OS.ObjRange Is Nothing Then
2468          X(1, 1) = 0
          ' If objective cell is error, report this directly to NOMAD. Attempting to manipulate it can cause errors
2469      ElseIf VarType(OS.ObjRange.Value2) = vbError Then
2470          X(1, 1) = OS.ObjRange.Value2
          'If objective sense is maximise then multiply by minus 1
2471      ElseIf OS.ObjectiveSense = MaximiseObjective Then
2472          If OS.ObjRange.Value2 <> 0 Then
2473              X(1, 1) = -1 * OS.ObjRange.Value2 'objective value
2474          Else
2475              X(1, 1) = OS.ObjRange.Value2
2476          End If
          'Else if objective sense is minimise leave it
2477      ElseIf OS.ObjectiveSense = MinimiseObjective Then
2478          X(1, 1) = OS.ObjRange.Value2
2479      ElseIf OS.ObjectiveSense = TargetObjective Then
2480          X(1, 1) = Abs(OS.ObjRange.Value2 - OS.ObjectiveTargetValue)
2481      End If
2483      k = 1 'keep a count of what constraint its up to not including bounds
          Dim row As Long, constraint As Long
2484      row = 1
          Dim CurrentLHSValues As Variant
          Dim CurrentRHSValues As Variant
2485      For constraint = 1 To OS.NumConstraints
              ' Check to see what is different and add rows to sparsea
2486          If Not OS.LHSRange(constraint) Is Nothing Then ' skip Binary and Integer constraints
                  ' Get current value(s) for LHS and RHS of this constraint off the sheet. LHS is always an array (even if 1x1)
2487              OS.GetCurrentConstraintValues constraint, CurrentLHSValues, CurrentRHSValues
2488              If OS.LHSType(constraint) = SolverInputType.MultiCellRange Then
2489                  For i = 1 To UBound(OS.LHSOriginalValues(constraint), 1)
2490                      For j = 1 To UBound(OS.LHSOriginalValues(constraint), 2)
2491                          If OS.RowSetsBound(row) = False Then
2492                              If OS.RHSType(constraint) <> SolverInputType.MultiCellRange Then
2493                                  SetConstraintValue X, k, CurrentRHSValues, CurrentLHSValues(i, j), OS.Relation(row)
2494                              ElseIf UBound(OS.LHSOriginalValues(constraint), 1) = UBound(OS.RHSOriginalValues(constraint), 1) Then
2495                                  SetConstraintValue X, k, CurrentRHSValues(i, j), CurrentLHSValues(i, j), OS.Relation(row)
2496                              Else
2497                                  SetConstraintValueMismatchedDims X, k, CurrentRHSValues, CurrentLHSValues, OS.Relation(row), i, j
2498                              End If
2499                          End If
2500                          row = row + 1
2501                      Next j
2502                  Next i
2503              Else
2504                  If OS.RowSetsBound(row) = False Then
2505                      SetConstraintValue X, k, CurrentRHSValues, CurrentLHSValues(1, 1), OS.Relation(row)
2506                  End If
2507                  row = row + 1
2508              End If
2509          End If
2510      Next constraint
          
          'Get back new objective and difference between LHS and RHS values
2511      NOMAD_GetValues = X
End Function

Sub NOMAD_RecalculateValues()
7129      If Not ForceCalculate("Warning: The worksheet calculation did not complete, and so the iteration may not be calculated correctly. Would you like to retry?") Then Exit Sub
End Sub

Function NOMAD_GetNumVariables() As Variant
7130      NOMAD_GetNumVariables = OS.AdjustableCells.Count
End Function

Function NOMAD_GetNumConstraints() As Variant
          'The number of constraints is actually the number of Objectives + Number of Constraints
          'Note: Bounds do not count as constraints and equalities count as 2 constraints
          Dim row As Long
          Dim X(0 To 1) As Double
2540      X(0) = 1
2541      For row = 1 To OS.NumRows
2542          If OS.RowSetsBound(row) = False Then X(0) = X(0) + 1
2543          If OS.Relation(row) = RelationEQ Then X(0) = X(0) + 1
2544      Next row

          'Number of objectives - NOMAD can do bi-objective
          'Note: Currently OpenSolver can only do single objectives- will need to set up multi objectives yourself
2545      X(1) = 1 'number of objectives

7131      NOMAD_GetNumConstraints = X
End Function

Function NOMAD_GetVariableData() As Variant
7132      Dim numVars As Double
2547      numVars = NOMAD_GetNumVariables

          Dim X() As Double
2548      ReDim X(0 To 4 * numVars - 1)

          ' Get bounds
          Dim i As Long, j As Long
2549      For i = 0 To numVars - 1
2550          If OS.AssumeNonNegativeVars Then
2551              X(2 * i) = 0
2552          Else
2553              X(2 * i) = -10000000000000#
2554          End If
2555          X(2 * i + 1) = 10000000000000#
2556      Next i

          ' Apply bounds
          Dim var As Long, c As Range
          var = 0
2562      For Each c In OS.AdjustableCells
2563          If TestKeyExists(OS.VarLowerBounds, c.Address) Then
2564              X(2 * var) = OS.VarLowerBounds(c.Address)
2565          End If
              If TestKeyExists(OS.VarUpperBounds, c.Address) Then
                  X(2 * var + 1) = OS.VarUpperBounds(c.Address)
              End If
              var = var + 1
2566      Next c

          'Get the starting point
          'Takes the points on the sheet and forces them between the bounds
          Dim CellValue As Double
          j = 0
2567      For Each c In OS.AdjustableCells
              CellValue = c.value
2568          If CellValue < X(2 * j) Then
2569              X(j + 2 * numVars) = X(2 * j)
2570          ElseIf CellValue > X(2 * j + 1) Then
2571              X(j + 2 * numVars) = X(2 * j + 1)
2572          Else
2573              X(j + 2 * numVars) = CellValue
2574          End If
              j = j + 1
2575      Next c
          
          'Get the variable type(real, int or bin)
2576      For i = 1 To numVars
          'initialise all variables as continuous
2577          X(i - 1 + 3 * numVars) = 1
2578      Next i
          Dim counter As Long, types As Variant
2579      counter = 2
2580      For Each types In Array(OS.IntegerCellsRange, OS.BinaryCellsRange)
2581          If Not types Is Nothing Then
2582              For Each c In types
2583                  For i = 1 To numVars
2584                      If OS.VarNames(i) = c.Address(RowAbsolute:=False, ColumnAbsolute:=False) Then
2585                          X(i - 1 + 3 * numVars) = counter
2586                          If Not OS.SolveRelaxation Then
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
          
2611      NOMAD_GetVariableData = X
End Function

Function NOMAD_GetOptionData() As Variant
          IterationCount = 0

          Dim SolverParameters As Dictionary
          Set SolverParameters = OS.SolverParameters
          ' Add extra values that depend on precision
          If SolverParameters.Exists(OS.Solver.PrecisionName) Then
              SolverParameters.Add Key:="H_MIN", Item:=SolverParameters.Item(OS.Solver.PrecisionName)
          End If
          
          Dim X() As Variant
          ReDim X(1 To SolverParameters.Count, 1 To 2)

          Dim i As Long, Key As Variant
          i = 1
          For Each Key In SolverParameters.Keys
              X(i, 1) = Key & " " & StrExNoPlus(SolverParameters.Item(Key))
              X(i, 2) = Len(X(i, 1))
              i = i + 1
          Next Key
          
          NOMAD_GetOptionData = X
End Function

Function NOMAD_ShowCancelDialog() As Variant
          Dim Response As VbMsgBoxResult
          Response = ShowEscapeCancelMessage()
          NOMAD_ShowCancelDialog = (Response = vbYes)
End Function

Private Sub SetConstraintValue(ByRef ConstraintValues As Variant, ByRef k As Long, RHSValue As Variant, LHSValue As Variant, RelationType As Long)
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

Private Sub SetConstraintValueMismatchedDims(ByRef ConstraintValues As Variant, ByRef k As Long, RHSValues As Variant, LHSValues As Variant, RelationType As Long, i As Long, j As Long)
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

Private Function DifferenceOrError(Value1 As Variant, Value2 As Variant) As Variant
2532            If VarType(Value1) = vbError Then
2533                DifferenceOrError = Value1
2534            ElseIf VarType(Value2) = vbError Then
2535                DifferenceOrError = Value2
2536            Else
                    On Error GoTo ErrorHandler
2537                DifferenceOrError = Value1 - Value2
2538            End If
                Exit Function
                
ErrorHandler:
                DifferenceOrError = Value1
End Function
