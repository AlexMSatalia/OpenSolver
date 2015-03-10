VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CQuickSolve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private NumRows As Long
Private NumParams As Long
Private OriginalParamValues() As Double
Private OriginalRHS() As Double

Private SparseA() As CIndexedCoeffs

Private CurrentParameterRange As Range

Function AnalyseParameters(OpenSolver As COpenSolver, sheet As Worksheet, _
                NumConstraints As Long, ByRef RHS() As Double, _
                ByRef LHSRange() As Range, ByRef RHSRange() As Range, _
                ByRef RHSFormula() As String, _
                ByRef RHSType() As SolverInputType, ByRef LHSOriginalValues As Variant, ByRef RHSOriginalValues As Variant) As Boolean

          ' This builds a model of the form RHS = InitialRHS + A_param * (Param - InitialParam)
          ' where b = solver model constraint RHS values
          '       b_param = value of b when each parameter is 0
          '       A_param is a matrix
          '       Param is the vector of the current parameter values
          ' Returns true if successful
          ' Note: Will update arguments LHSValue and RHSValue to the current values of these cells on the spreadsheet
          ' This remembers the parameter range for which we have built the model for checking later
          
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

2619      Set CurrentParameterRange = Nothing
          
          ' Remember the current number of rows and original RHS values
          Dim i As Long, oldCalculationMode As Long
          
2620      NumRows = OpenSolver.GetNumRows
          
2621      oldCalculationMode = Application.Calculation

2622      If NumRows > 0 Then ReDim OriginalRHS(NumRows) As Double
2623      For i = 1 To NumRows
2624          OriginalRHS(i) = RHS(i)
2625      Next i
          
          ' Find the Parameter range, and get the numer of parameters
          Dim ParamRange As Range
2626      If Not OpenSolverModule.CheckModelHasParameterRange Then GoTo ExitFunction   ' Error message already shown
2629      Set ParamRange = OpenSolverModule.RemoveRangeOverlap(GetQuickSolveParameters(sheet:=sheet))
          
          Dim BadCell As Range
2630      If Not OpenSolverModule.CheckRangeContainsNoAmbiguousMergedCells(ParamRange, BadCell) Then
2631          Err.Raise OpenSolver_BuildError, Description:="QuickSolve Initialisation Failed: The parameter " & BadCell.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & " is inaccessible as it is within the merged range " & BadCell.MergeArea.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "."
2633      End If
          
2634      NumParams = ParamRange.Count
          
          ' Remember the original parameter values
          ' Warning: Do not use ParamRange(i) as this fails with more than 1 area; using "for each" seems to work
          ' TODO: THis is very slow; speed it up
          Dim c As Range
2635      If NumParams > 0 Then ReDim OriginalParamValues(NumParams) As Double
2636      i = 1
2637      For Each c In ParamRange
2638          OriginalParamValues(i) = c.Value2
2639          i = i + 1
2640      Next c
          
          ' Get current values for the LHS cell and RHS cell of the Excel constraints
          ' Note: If LHSIsRange(i)=false, then LHSValue(i) already contains the fixed constant value given in the constraint
          
          ' Look at how the LHSRange and RHSRange values now change as we perturb the parameters, and hence deduce how the OpenSolver model's RHS values alter
2641      Application.Calculation = xlCalculationManual
2642      Application.ScreenUpdating = False
          
          Dim param As Long, row As Long, NonZeros As Long, constraint As Long
          Dim CurrentLHSValues  As Variant, CurrentRHSValues As Variant
2643      param = 0
          
2644      If NumRows > 0 And NumParams > 0 Then ReDim SparseA(NumRows)
2645      For row = 1 To NumRows
2646          Set SparseA(row) = New CIndexedCoeffs
2647      Next row
         
2648      For Each c In ParamRange
2649          If c.HasArray Then
2650              Err.Raise OpenSolver_BuildError, Description:="QuickSolve Initialisation Failed: OpenSolver cannot perform quick solves for parameters that contain array formulae."
2652          End If

2653          param = param + 1
2654          UpdateStatusBar "OpenSolver: Analysing Parameters... " & param & "/" & NumParams & " parameters"
              
              Dim OldFormula As String
2655          On Error Resume Next
2656          OldFormula = c.Formula
2657          If Err.Number <> 0 Then
2658              Err.Raise OpenSolver_BuildError, Description:="QuickSolve Initialisation Failed: An error (" & Err.Description & ") occurred while changing the value in parameter cell " & c.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "."
2660          End If
              ' Change the parameter value, and see how the RHS's change
2661          On Error GoTo ErrorHandler
2662          c.Value2 = OriginalParamValues(param) + 1
2663          Application.Calculate
2664          row = 1
              Dim j As Long, coeff As Double
2665          For constraint = 1 To NumConstraints
                  ' Check to see what is different and add rows to sparsea
2666              If Not LHSRange(constraint) Is Nothing Then ' skip Binary and Integer constraints
                      Dim OriginalLHSValues As Variant
                      Dim OriginalRHSValues As Variant
2667                  OriginalLHSValues = LHSOriginalValues(constraint)
2668                  OriginalRHSValues = RHSOriginalValues(constraint)
                      ' Get current value(s) for LHS and RHS of this constraint off the sheet. LHS is always an array (even if 1x1)
2669                  OpenSolver.GetCurrentConstraintValues constraint, CurrentLHSValues, CurrentRHSValues
                      'TODO Put in better error trapping for errors that occur on the s/sheet
2670                  For i = 1 To UBound(OriginalLHSValues, 1) ' number of rows
2671                      For j = 1 To UBound(OriginalLHSValues, 2) ' number of columns
2672                          coeff = -(CurrentLHSValues(i, j) - OriginalLHSValues(i, j))
                              If RHSType(constraint) = SolverInputType.MultiCellRange Then
2673                              coeff = coeff + (CurrentRHSValues(i, j) - OriginalRHSValues(i, j))
2678                          ElseIf RHSType(constraint) = SolverInputType.SingleCellRange Or RHSType(constraint) = SolverInputType.Formula Then
2679                              coeff = coeff + (CurrentRHSValues - OriginalRHSValues)
2680                          End If
2685                          If Not IsZero(coeff) Then
2687                              SparseA(row).Add param, coeff
2688                              NonZeros = NonZeros + 1
2689                          End If
2691                          row = row + 1
2692                      Next j
2693                  Next i
2694              End If
2695          Next constraint
2696          c.Formula = OldFormula
2697      Next c
          
2698      If NonZeros = 0 Then
2699          Err.Raise OpenSolver_BuildError, Description:="QuickSolve Initialisation Failed: The model's constraints do not depend on the QuickSolve parameter cells; please try again. (The QuickSolve process is used for repeatedly solving a model after making changes to the parameter cells, where these changes result in changes to the model's constraints.)"
2701      End If
          
2702      Set CurrentParameterRange = ParamRange
2703      AnalyseParameters = True
          
ExitFunction:
2704      Application.StatusBar = False ' Resume normal status bar behaviour
2705      Application.ScreenUpdating = True
2706      Application.Calculation = oldCalculationMode
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("CQuickSolve", "AnalyseParameters") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Public Function UpdateModelRHS(sheet As Worksheet, ByRef RHS() As Double) As Boolean
' Given the new Parameter values on the worksheet, calculate a new set of model RHS values
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          ' Find the Parameter range on the given (current) sheet, and check that this is the range we are set up for
          Dim ParamRange As Range
2719      Set ParamRange = GetQuickSolveParameters(sheet:=sheet)
2720      If ParamRange Is Nothing Then
2721          Err.Raise OpenSolver_BuildError, Description:="No Parameter range could be found on the worksheet."
2724      End If
2726      If (CurrentParameterRange.Address <> ParamRange.Address) Or (CurrentParameterRange.Worksheet.Name <> ParamRange.Worksheet.Name) Then
2727          Err.Raise OpenSolver_BuildError, Description:="QuickSolve Failed: You need to initialize the QuickSolve engine for the current parameters."
2730      End If
2731      If NumParams <> ParamRange.Count Then
2732          Err.Raise OpenSolver_BuildError, Description:="QuickSolve Failed: The number of values in the Parameter range has changed, and so the model update can not be completed."
2735      End If
          
          Dim ParamChanges() As Double
2736      If NumParams > 0 Then ReDim ParamChanges(NumParams) As Double
          Dim param As Long, c As Range
2737      param = 1
2738      For Each c In ParamRange
2739          ParamChanges(param) = c.Value2 - OriginalParamValues(param)
2740          param = param + 1
2741      Next c
          
          Dim row As Long, value As Double, i As Long, coeff As Double
2742      For row = 1 To NumRows
2743          value = OriginalRHS(row)

2744          With SparseA(row)
2745              For i = 1 To .Count
2746                  param = .Index(i)
2747                  coeff = .Coefficient(i)
2748                  value = value + coeff * ParamChanges(param)
2749              Next i
2750          End With
2751          RHS(row) = value
2752      Next row
          
2753      UpdateModelRHS = True

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("CQuickSolve", "UpdateModelRHS") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function
