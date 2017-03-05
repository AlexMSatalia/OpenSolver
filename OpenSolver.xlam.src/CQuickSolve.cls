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

Function AnalyseParameters(s As COpenSolver, ParamRange As Range) As Boolean

          ' This builds a model of the form RHS = InitialRHS + A_param * (Param - InitialParam)
          ' where b = solver model constraint RHS values
          '       b_param = value of b when each parameter is 0
          '       A_param is a matrix
          '       Param is the vector of the current parameter values
          ' Returns true if successful
          ' Note: Will update arguments LHSValue and RHSValue to the current values of these cells on the spreadsheet
          ' This remembers the parameter range for which we have built the model for checking later
          
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         Set CurrentParameterRange = Nothing
          
          ' Remember the current number of rows and original RHS values
          Dim i As Long, oldCalculationMode As Long
          
4         NumRows = s.NumRows
          
5         oldCalculationMode = Application.Calculation

6         If NumRows > 0 Then ReDim OriginalRHS(NumRows) As Double
7         For i = 1 To NumRows
8             OriginalRHS(i) = s.RHS(i)
9         Next i
          
          Dim BadCell As Range
10        If Not CheckRangeContainsNoAmbiguousMergedCells(ParamRange, BadCell) Then
11            RaiseUserError "QuickSolve Initialisation Failed: The parameter " & BadCell.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & " is inaccessible as it is within the merged range " & BadCell.MergeArea.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "."
12        End If
          
13        NumParams = ParamRange.Count
          
          ' Remember the original parameter values
          ' Warning: Do not use ParamRange(i) as this fails with more than 1 area; using "for each" seems to work
          ' TODO: THis is very slow; speed it up
          Dim c As Range
14        If NumParams > 0 Then ReDim OriginalParamValues(NumParams) As Double
15        i = 1
16        For Each c In ParamRange
17            OriginalParamValues(i) = c.Value2
18            i = i + 1
19        Next c
          
          ' Get current values for the LHS cell and RHS cell of the Excel constraints
          ' Note: If LHSIsRange(i)=false, then LHSValue(i) already contains the fixed constant value given in the constraint
          
          ' Look at how the LHSRange and RHSRange values now change as we perturb the parameters, and hence deduce how the OpenSolver model's RHS values alter
20        Application.Calculation = xlCalculationManual
          Dim ScreenStatus As Boolean
21        ScreenStatus = Application.ScreenUpdating
22        Application.ScreenUpdating = False
          
          Dim param As Long, row As Long, NonZeros As Long, constraint As Long
          Dim CurrentLHSValues As Variant, CurrentRHSValues As Variant
23        param = 0
          
24        If NumRows > 0 And NumParams > 0 Then ReDim SparseA(NumRows)
25        For row = 1 To NumRows
26            Set SparseA(row) = New CIndexedCoeffs
27        Next row
          
          Dim LHSOriginalValues() As Variant, RHSOriginalValues() As Variant
28        ReDim LHSOriginalValues(s.NumConstraints)
29        ReDim RHSOriginalValues(s.NumConstraints)
30        For constraint = 1 To s.NumConstraints
31            If Not s.LHSRange(constraint) Is Nothing Then
32                s.GetCurrentConstraintValues constraint, LHSOriginalValues(constraint), RHSOriginalValues(constraint)
33            End If
34        Next constraint
          
35        For Each c In ParamRange
36            If c.HasArray Then
37                RaiseUserError "QuickSolve Initialisation Failed: OpenSolver cannot perform quick solves for parameters that contain array formulae."
38            End If

39            param = param + 1
40            UpdateStatusBar "OpenSolver: Analysing Parameters... " & param & "/" & NumParams & " parameters"
              
              Dim OldFormula As String
41            On Error Resume Next
42            OldFormula = c.Formula
43            If Err.Number <> 0 Then
44                RaiseUserError "QuickSolve Initialisation Failed: An error (" & Err.Description & ") occurred while changing the value in parameter cell " & c.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & "."
45            End If
              ' Change the parameter value, and see how the RHS's change
46            On Error GoTo ErrorHandler
47            c.Value2 = OriginalParamValues(param) + 1
48            Application.Calculate
49            row = 1
              Dim j As Long, coeff As Double
50            For constraint = 1 To s.NumConstraints
                  ' Check to see what is different and add rows to sparsea
51                If Not s.LHSRange(constraint) Is Nothing Then ' skip Binary and Integer constraints
                      Dim OriginalLHSValues As Variant
                      Dim OriginalRHSValues As Variant
52                    OriginalLHSValues = LHSOriginalValues(constraint)
53                    OriginalRHSValues = RHSOriginalValues(constraint)
                      ' Get current value(s) for LHS and RHS of this constraint off the sheet. LHS is always an array (even if 1x1)
54                    s.GetCurrentConstraintValues constraint, CurrentLHSValues, CurrentRHSValues
                      'TODO Put in better error trapping for errors that occur on the s/sheet
55                    For i = 1 To UBound(OriginalLHSValues, 1) ' number of rows
56                        For j = 1 To UBound(OriginalLHSValues, 2) ' number of columns
57                            coeff = -(CurrentLHSValues(i, j) - OriginalLHSValues(i, j))
58                            If s.RHSType(constraint) = SolverInputType.MultiCellRange Then
59                                coeff = coeff + (CurrentRHSValues(i, j) - OriginalRHSValues(i, j))
60                            ElseIf s.RHSType(constraint) = SolverInputType.SingleCellRange Or s.RHSType(constraint) = SolverInputType.Formula Then
61                                coeff = coeff + (CurrentRHSValues - OriginalRHSValues)
62                            End If
63                            If Not IsZero(coeff) Then
64                                SparseA(row).Add param, coeff
65                                NonZeros = NonZeros + 1
66                            End If
67                            row = row + 1
68                        Next j
69                    Next i
70                End If
71            Next constraint
72            c.Formula = OldFormula
73        Next c
          
74        If NonZeros = 0 Then
75            RaiseUserError "QuickSolve Initialisation Failed: The model's constraints do not depend on the QuickSolve parameter cells; please try again. (The QuickSolve process is used for repeatedly solving a model after making changes to the parameter cells, where these changes result in changes to the model's constraints.)"
76        End If
          
77        Set CurrentParameterRange = ParamRange
78        AnalyseParameters = True
          
ExitFunction:
79        Application.StatusBar = False ' Resume normal status bar behaviour
80        Application.ScreenUpdating = ScreenStatus
81        Application.Calculation = oldCalculationMode
82        If RaiseError Then RethrowError
83        Exit Function

ErrorHandler:
84        If Not ReportError("CQuickSolve", "AnalyseParameters") Then Resume
85        RaiseError = True
86        GoTo ExitFunction
End Function

Public Function UpdateModelRHS(sheet As Worksheet, ByRef RHS() As Double) As Boolean
      ' Given the new Parameter values on the worksheet, calculate a new set of model RHS values
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          ' Find the Parameter range on the given (current) sheet, and check that this is the range we are set up for
          Dim ParamRange As Range
3         Set ParamRange = GetQuickSolveParameters(sheet, Validate:=True)

4         If (CurrentParameterRange.Address <> ParamRange.Address) Or _
             (CurrentParameterRange.Worksheet.Name <> ParamRange.Worksheet.Name) Or _
             (CurrentParameterRange.Worksheet.Parent.Name <> ParamRange.Worksheet.Parent.Name) Then
5             RaiseUserError "QuickSolve Failed: You need to initialize the QuickSolve engine for the current parameters."
6         End If
7         If NumParams <> ParamRange.Count Then
8             RaiseUserError "QuickSolve Failed: The number of values in the Parameter range has changed, and so the model update can not be completed."
9         End If
          
          Dim ParamChanges() As Double
10        If NumParams > 0 Then ReDim ParamChanges(NumParams) As Double
          Dim param As Long, c As Range
11        param = 1
12        For Each c In ParamRange
13            ParamChanges(param) = c.Value2 - OriginalParamValues(param)
14            param = param + 1
15        Next c
          
          Dim row As Long, value As Double, i As Long, coeff As Double
16        For row = 1 To NumRows
17            value = OriginalRHS(row)

18            With SparseA(row)
19                For i = 1 To .Count
20                    param = .Index(i)
21                    coeff = .Coefficient(i)
22                    value = value + coeff * ParamChanges(param)
23                Next i
24            End With
25            RHS(row) = value
26        Next row
          
27        UpdateModelRHS = True

ExitFunction:
28        If RaiseError Then RethrowError
29        Exit Function

ErrorHandler:
30        If Not ReportError("CQuickSolve", "UpdateModelRHS") Then Resume
31        RaiseError = True
32        GoTo ExitFunction
End Function
