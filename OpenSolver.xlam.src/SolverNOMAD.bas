Attribute VB_Name = "SolverNOMAD"
Option Explicit

Public OS As COpenSolver
Dim IterationCount As Long

Dim numVars As Long

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
2547      numVars = NOMAD_GetNumVariables

          Dim X() As Double
2548      ReDim X(0 To 4 * numVars - 1)

          Dim DefaultLowerBound As Double, DefaultUpperBound As Double
          DefaultLowerBound = IIf(OS.AssumeNonNegativeVars, 0, -10000000000000#)
          DefaultUpperBound = 10000000000000#

          Dim i As Long, c As Range
          i = 0
          For Each c In OS.AdjustableCells
              ' Set the default bounds
              SetLowerBound X, i, DefaultLowerBound
              SetUpperBound X, i, DefaultUpperBound
              
              ' Set any specified bounds (overwriting defaults)
              If TestKeyExists(OS.VarLowerBounds, c.Address) Then
                  SetLowerBound X, i, OS.VarLowerBounds(c.Address)
              End If
              If TestKeyExists(OS.VarUpperBounds, c.Address) Then
                  SetUpperBound X, i, OS.VarUpperBounds(c.Address)
              End If
              
              Dim LowerBound As Double, UpperBound As Double
              LowerBound = GetLowerBound(X, i)
              UpperBound = GetUpperBound(X, i)
          
              ' Get the starting point
              SetStartingPoint X, i, c.value
          
              ' Initialise all variables as continuous
              SetVarType X, i, VariableType.VarContinuous
              
              ' Get the variable type (int or bin)
2584          If OS.SolveRelaxation Then
                  ' Set Binary vars to have bounds of 0 and 1 and start at 0
                  If TestIntersect(c, OS.BinaryCellsRange) Then
                      SetLowerBound X, i, 0
                      SetUpperBound X, i, 1
                      SetStartingPoint X, i, 0
                  End If
              Else
                  Dim Integral As Boolean
                  Integral = True
                  If TestIntersect(c, OS.BinaryCellsRange) Then
                      SetVarType X, i, VariableType.VarBinary
                  ElseIf TestIntersect(c, OS.IntegerCellsRange) Then
                      SetVarType X, i, VariableType.VarInteger
                  Else
                      Integral = False
                  End If
                  
                  If Integral Then
                      'Make bounds on integer and binary constraints integer
2587                  If LowerBound > 0 Then
2588                      LowerBound = Application.WorksheetFunction.RoundUp(LowerBound, 0)
2589                  Else
2590                      LowerBound = Application.WorksheetFunction.RoundDown(LowerBound, 0)
2591                  End If
                      SetLowerBound X, i, LowerBound

2592                  If UpperBound > 0 Then
2593                      UpperBound = Application.WorksheetFunction.RoundDown(UpperBound, 0)
2594                  Else
2595                      UpperBound = Application.WorksheetFunction.RoundUp(UpperBound, 0)
2596                  End If
                      SetUpperBound X, i, UpperBound
                      
                      'Make starting positions on integer and binary constraints integer
                      SetStartingPoint X, i, Round(GetStartingPoint(X, i))
                  End If
              End If
              
              ' Force starting point between the bounds
              Dim StartingPoint As Double
              StartingPoint = GetStartingPoint(X, i)
              If StartingPoint < LowerBound Then
                  StartingPoint = LowerBound
              ElseIf StartingPoint > UpperBound Then
                  StartingPoint = UpperBound
              End If
              SetStartingPoint X, i, StartingPoint
              
              i = i + 1
          Next c
          
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

Private Function GetLowerBound(X As Variant, i As Long) As Double
    GetLowerBound = X(i * 2)
End Function

Private Sub SetLowerBound(ByRef X As Variant, i As Long, value As Double)
    X(i * 2) = value
End Sub

Private Function GetUpperBound(X As Variant, i As Long) As Double
    GetUpperBound = X(i * 2 + 1)
End Function

Private Sub SetUpperBound(ByRef X As Variant, i As Long, value As Double)
    X(i * 2 + 1) = value
End Sub

Private Function GetStartingPoint(X As Variant, i As Long) As Double
    GetStartingPoint = X(numVars * 2 + i)
End Function

Private Sub SetStartingPoint(X As Variant, i As Long, value As Double)
    X(numVars * 2 + i) = value
End Sub

Private Function GetVarType(X As Variant, i As Long) As Double
    GetVarType = X(numVars * 3 + i)
End Function

Private Sub SetVarType(X As Variant, i As Long, value As Double)
    X(numVars * 3 + i) = value
End Sub

