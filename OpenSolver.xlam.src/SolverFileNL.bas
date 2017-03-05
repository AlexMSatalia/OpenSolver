Attribute VB_Name = "SolverFileNL"
Option Explicit

Public Const NLModelFileName As String = "model.nl"

' This module is for writing .nl files that describe the model and solving these
' For more info on .nl file format see:
' http://citeseerx.ist.psu.edu/viewdoc/summary?doi=10.1.1.60.9659
' http://www.ampl.com/REFS/hooking2.pdf

Dim m As CModelParsed
Dim s As COpenSolver
Dim SolveScriptPathName As String

Dim problem_name As String

Dim WriteComments As Boolean    ' Whether .nl file should include comments
Const CommentIndent = 4         ' Tracks the level of indenting in comments on nl output
Public Const CommentSpacing = 28       ' The column number at which nl comments begin

' ==========================================================================
' ASL variables
' These are variables used in the .NL file that are also used by the AMPL Solver Library
' We use the same names as the ASL for consistency
' ASL header definitions available at http://www.netlib.org/ampl/solvers/asl.h

Dim n_var As Long            ' Number of variables
Dim n_con As Long            ' Number of constraints
Dim n_obj As Long            ' Number of objectives
Dim nranges As Long          ' Number of range constraints
Dim n_eqn_ As Long           ' Number of equality constraints
Dim n_lcon As Long           ' Number of logical constraints

Dim nlc As Long              ' Number of non-linear constraints
Dim nlo As Long              ' Number of non-linear objectives

Dim nlnc As Long             ' Number of non-linear network constraints
Dim lnc As Long              ' Number of linear network constraints

Dim nlvc As Long             ' Number of variables appearing non-linearly in constraints
Dim nlvo As Long             ' Number of variables appearing non-linearly in objectives
Dim nlvb As Long             ' Number of variables appearing non-linearly in both constraints and objectives

Dim nwv_ As Long             ' Number of linear network variables
Dim nfunc_ As Long           ' Number of user-defined functions
Dim arith As Long            ' Not sure what this does, keep at 0. This flag may indicate whether little- or big-endian arithmetic was used to write the file.
Dim flags As Long            ' 1 = want output suffixes

Dim nbv As Long              ' Number of binary variables
Dim niv As Long              ' Number of other integer variables
Dim nlvbi As Long            ' Number of integer variables appearing non-linearly in both constraints and objectives
Dim nlvci As Long            ' Number of integer variables appearing non-linearly just in constraints
Dim nlvoi As Long            ' Number of integer variables appearing non-linearly just in objectives

Dim nzc As Long              ' Number of non-zeros in the Jacobian matrix
Dim nzo As Long              ' Number of non-zeros in objective gradients

Dim maxrownamelen_ As Long   ' Length of longest constraint name
Dim maxcolnamelen_ As Long   ' Length of longest variable name

' Common expressions (from "defined" vars) - whatever that means!? We just set them to zero as we don't use defined vars
Dim comb As Long
Dim comc As Long
Dim como As Long
Dim comc1 As Long
Dim como1 As Long

' End ASL variables
' ===========================================================================

' The following variables are used for storing the NL model
' Definitions for some of the terms used:
'   - parsed variable index/order:      the order that variables are read into the model. There are numActualVars adjustable cells, followed by numFakeVars formulae variables
'   - .nl variable index/order:         the order that variables are arranged in the .nl output. This follows a strict specification (see Sub MakeVariableMap)
'   - parsed constraint index/order:    the order that constraints are read into the model. There are numActualCons real constraints followed by numFakeCons formulae constraints
'   - .nl constraint index/order:       the order that constraints are arranged in the .nl output. This follows a strict specification (see Sub MakeConstraintMap)
'   - parsed objective index/order:     there is currently only a single objective supported, so this is irrelevant

Dim VariableMap As Dictionary         ' A map from variable name (e.g. Test1_D4) to .nl variable index (0 to n_var - 1)
Dim VariableMapRev() As String        ' A map from .nl variable index (0 to n_var - 1) to variable name (e.g. Test1_D4)
Public VariableIndex As Dictionary    ' A map from variable name (e.g. Test1_D4) to parsed variable index (1 to n_var)

Dim ConstraintMapRev() As String     ' A map from .nl constraint index (0 to n_con - 1) to constraint name (e.g. c1_Test1_D4)

Dim NonLinearConstraintTrees() As ExpressionTree  ' Array of size n_con containing all non-linear constraint ExpressionTrees stored in parsed constraint order
Dim NonLinearObjectiveTrees() As ExpressionTree   ' Array of size n_con containing all non-linear objective ExpressionTrees stored in parsed objective order

Dim NonLinearVars() As Boolean          ' Array of size n_var indicating whether each variable appears non-linearly in the model
Dim NonLinearConstraints() As Boolean   ' Array of size n_con indicating whether each constraint has non-linear elements

Dim ConstraintIndexToTreeIndex() As Long         ' Array of size n_con storing the parsed constraint index for each .nl constraint index
Dim VariableNLIndexToCollectionIndex() As Long   ' Array of size n_var storing the parsed variable index for each .nl variable index
Dim VariableCollectionIndexToNLIndex() As Long   ' Array of size n_var storing the .nl variable index for each parsed variable index

Dim LinearConstraints() As Dictionary     ' Array of size n_con containing the Dictionaries for each constraint stored in parsed constraint order
Dim LinearConstants() As Double           ' Array of size n_con containing the constant (a Double) for each constraint stored in parsed constraint order
Dim LinearObjectives() As Dictionary      ' Array of size n_obj containing the Dictionaries for each objective stored in parsed objective order

Dim InitialVariableValues() As Double ' Array of size n_var containing the intital values for each variable in OpenSolver index order

Dim ConstraintRelations() As RelationConsts   ' Array of size n_con containing the RelationConst for each constraint stored in parsed constraint order

Dim ObjectiveCells() As String                ' Array of size n_obj containing the objective cells for each objective stored in parsed objective order
Dim ObjectiveSenses() As ObjectiveSenseType   ' Array of size n_obj containing the objective sense for each objective stored in parsed objective order

Dim numActualVars As Long    ' The number of actual variables in the parsed model (the adjustable cells in the Solver model)
Dim numFakeVars As Long      ' The number of "fake" variables in the parsed model (variables that arise from parsing the formulae in the Solver model)
Dim numActualCons As Long    ' The number of actual constraints in the parsed model (the constraints defined in the Solver model)
Dim numFakeCons As Long      ' The number of "fake" constraints in the parsed model (the constraints that arise from parsing the formulae in the Solver model)
Dim numActualEqs As Long     ' The number of actual equality constraints in the parsed model
Dim numActualRanges As Long  ' The number of actual inequality constraints in the parsed model

Dim NonZeroConstraintCount() As Long ' An array of size n_var counting the number of times that each variable (in parsed variable order) appears in the linear constraints

Dim BinaryVars() As Boolean

' Public accessor for VariableCollectionIndexToNLIndex
Public Property Get GetVariableNLIndex(Index As Long) As Long
1         GetVariableNLIndex = VariableCollectionIndexToNLIndex(Index)
End Property

Function GetNLModelFilePath(ByRef Path As String) As Boolean
1               GetNLModelFilePath = GetTempFilePath(NLModelFileName, Path)
End Function

' Creates .nl file and solves model
Function WriteNLFile_Parsed(OpenSolver As COpenSolver, ModelFilePathName As String, Optional ShouldWriteComments As Boolean = True)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
3         Application.EnableCancelKey = xlErrorHandler

4         Set s = OpenSolver
5         Set m = s.ParsedModel

          Dim LocalExecSolver As ISolverLocalExec
6         Set LocalExecSolver = s.Solver
          
7         WriteComments = ShouldWriteComments
          
          ' =============================================================
          ' Process model for .nl output
          ' All Module-level variables required for .nl output should be set in this step
          ' No modification to these variables should be done while writing the .nl file
          ' =============================================================
          
8         InitialiseModelStats s
          
9         If n_obj = 0 And n_con = 0 Then
10            RaiseUserError "The model has no constraints that depend on the adjustable cells, and has no objective. There is nothing for the solver to do."
11        End If
          
12        CreateVariableIndex
          
13        ProcessFormulae
14        ProcessObjective
          
15        MakeVariableMap s.SolveRelaxation
16        MakeConstraintMap
          
          ' =============================================================
          ' Write output files
          ' =============================================================
          
          ' Create supplementary outputs
17        If WriteComments Then
18            OutputColFile
19            OutputRowFile
20        End If
          
          ' Write .nl file
21        Open ModelFilePathName For Output As #1
          
22        MakeHeader
23        MakeCBlocks
24        MakeOBlocks
          'MakeDBlock
25        If s.InitialSolutionIsValid Then MakeXBlock
26        MakeRBlock
27        MakeBBlock
28        MakeKBlock
29        MakeJBlocks
30        MakeGBlocks

ExitFunction:
31        Application.StatusBar = False
32        Close #1
33        If RaiseError Then RethrowError
34        Exit Function

ErrorHandler:
35        If Not ReportError("SolverFileNL", "SolveModelParsed_NL") Then Resume
36        RaiseError = True
37        GoTo ExitFunction
End Function

Private Sub InitialiseModelStats(s As COpenSolver)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          ' Number of actual variables is the number of adjustable cells in the Solver model
3         numActualVars = s.AdjustableCells.Count
          ' Number of fake variables is the number of formulae equations we have created
4         numFakeVars = m.Formulae.Count
          ' Number of actual constraints is the number of constraints in the Solver model
5         numActualCons = m.LHSKeys.Count
          ' Number of fake constraints is the number of formulae equations we have created
6         numFakeCons = m.Formulae.Count
          
          ' Count the actual constraints that are equalities and ranges (2-sided inequalities)
          Dim i As Long
7         numActualEqs = 0
8         numActualRanges = 0
9         For i = 1 To numActualCons
10            UpdateStatusBar "OpenSolver: Creating .nl file. Counting constraints: " & i & "/" & numActualCons & ". "
              
11            If m.RELs(i) = RelationConsts.RelationEQ Then
12                numActualEqs = numActualEqs + 1
              ' ElseIf
              '     ' We don't currently support 2-sided inequalities, so there are no ranges
              '     numActualRanges = numActualRanges + 1
13            End If
14        Next i
          
          ' ===============================================================================
          ' Initialise the ASL variables - see definitions for explanation of each variable
          
          ' Model statistics for line #1
15        problem_name = "Sheet=" + s.sheet.Name
          
          ' Model statistics for line #2
16        n_var = numActualVars + numFakeVars
17        n_con = numActualCons + numFakeCons + IIf(s.ObjectiveSense = TargetObjective, 1, 0)
18        n_obj = 0
19        nranges = numActualRanges
20        n_eqn_ = numActualEqs + numFakeCons     ' All fake formulae constraints are equalities
21        n_lcon = 0
          
          ' Model statistics for line #3
22        nlc = 0
23        nlo = 0
          
          ' Model statistics for line #4
24        nlnc = 0
25        lnc = 0
          
          ' Model statistics for line #5
26        nlvc = 0
27        nlvo = 0
28        nlvb = 0
          
          ' Model statistics for line #6
29        nwv_ = 0
30        nfunc_ = 0
31        arith = 0
32        flags = 1  ' We want suffixes printed in the .sol file
          
          ' Model statistics for line #7
33        nbv = 0
34        niv = 0
35        nlvbi = 0
36        nlvci = 0
37        nlvoi = 0
          
          ' Model statistics for line #8
38        nzc = 0
39        nzo = 0
          
          ' Model statistics for line #9
40        maxrownamelen_ = 0
41        maxcolnamelen_ = 0
          
          ' Model statistics for line #10
42        comb = 0
43        comc = 0
44        como = 0
45        comc1 = 0
46        como1 = 0

ExitSub:
47        If RaiseError Then RethrowError
48        Exit Sub

ErrorHandler:
49        If Not ReportError("SolverFileNL", "InitialiseModelStats") Then Resume
50        RaiseError = True
51        GoTo ExitSub
End Sub

' Creates map from variable name (e.g. Test1_D4) to parsed variable index (1 to n_var)
Private Sub CreateVariableIndex()
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         Set VariableIndex = New Dictionary
4         ReDim InitialVariableValues(1 To n_var) As Double

          Dim c As Range, cellName As String, i As Long
          
          ' First read in actual vars
5         i = 1
6         For Each c In s.AdjustableCells
7             UpdateStatusBar "OpenSolver: Creating .nl file. Counting variables: " & i & "/" & numActualVars & ". "
              
8             cellName = ConvertCellToStandardName(c)
              
              ' Update variable maps
9             VariableIndex.Add Key:=cellName, Item:=i
              
              ' Update initial values
10            InitialVariableValues(i) = c.Value2
              
11            i = i + 1
12        Next c
          
          ' Next read in fake formulae vars
13        For i = 1 To numFakeVars
14            UpdateStatusBar "OpenSolver: Creating .nl file. Counting formulae variables: " & i & "/" & numFakeVars & ". "
              
15            cellName = m.Formulae(i).strAddress
              
              ' Update variable maps
16            VariableIndex.Add Key:=cellName, Item:=i + numActualVars
17        Next i

ExitSub:
18        If RaiseError Then RethrowError
19        Exit Sub

ErrorHandler:
20        If Not ReportError("SolverFileNL", "CreateVariableIndex") Then Resume
21        RaiseError = True
22        GoTo ExitSub
End Sub

' Creates maps from variable name (e.g. Test1_D4) to .nl variable index (0 to n_var - 1) and vice-versa, and
' maps from parsed variable index to .nl variable index and vice-versa
Private Sub MakeVariableMap(SolveRelaxation As Boolean)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          ' Create index of variable names in parsed variable order
          Dim CellNames As New Collection
          
          ' Actual variables
          Dim c As Range, cellName As String, i As Long
3         i = 1
4         For Each c In s.AdjustableCells
5             UpdateStatusBar "OpenSolver: Creating .nl file. Classifying variables: " & i & "/" & numActualVars & ". "
              
6             i = i + 1
7             cellName = ConvertCellToStandardName(c)
8             CellNames.Add cellName
9         Next c
          
          ' Formulae variables
10        For i = 1 To m.Formulae.Count
11            UpdateStatusBar "OpenSolver: Creating .nl file. Classifying formulae variables: " & i & "/" & numFakeVars & ". "
              
12            cellName = m.Formulae(i).strAddress
13            CellNames.Add cellName
14        Next i
          
          '==============================================
          ' Get binary and integer variables from model
          
          ' Get integer variables
          Dim IntegerVars() As Boolean
15        ReDim IntegerVars(n_var) As Boolean
16        If Not s.IntegerCellsRange Is Nothing Then
17            UpdateStatusBar "OpenSolver: Creating .nl file. Finding integer variables"
18            For Each c In s.IntegerCellsRange
19                cellName = ConvertCellToStandardName(c)
20                IntegerVars(VariableIndex.Item(cellName)) = True
21            Next c
22        End If
          
          ' Get binary variables
23        ReDim BinaryVars(n_var) As Boolean
24        If Not s.BinaryCellsRange Is Nothing Then
25            UpdateStatusBar "OpenSolver: Creating .nl file. Finding binary variables"
26            For Each c In s.BinaryCellsRange
27                cellName = ConvertCellToStandardName(c)
28                BinaryVars(VariableIndex.Item(cellName)) = True
29            Next c
30        End If
          
          ' ==============================================
          ' Divide variables into the required groups. Note that there is no non-linear binary
          '     non-linear continuous
          '     non-linear integer
          '     linear continuous
          '     linear binary
          '     linear integer
          Dim NonLinearContinuous As New Collection
          Dim NonLinearInteger As New Collection
          Dim LinearContinuous As New Collection
          Dim LinearBinary As New Collection
          Dim LinearInteger As New Collection
          
31        For i = 1 To n_var
32            UpdateStatusBar "OpenSolver: Creating .nl file. Sorting variables: " & i & "/" & n_var & ". "
              
33            If NonLinearVars(i) Then
34                If IntegerVars(i) Or BinaryVars(i) Then
35                    NonLinearInteger.Add i
36                Else
37                    NonLinearContinuous.Add i
38                End If
39            Else
40                If BinaryVars(i) Then
41                    LinearBinary.Add i
42                ElseIf IntegerVars(i) Then
43                    LinearInteger.Add i
44                Else
45                    LinearContinuous.Add i
46                End If
47            End If
48        Next i
          
          ' ==============================================
          ' Add variables to the variable map in the required order
49        ReDim VariableNLIndexToCollectionIndex(n_var) As Long
50        ReDim VariableCollectionIndexToNLIndex(n_var) As Long
51        Set VariableMap = New Dictionary
52        ReDim VariableMapRev(n_var) As String
          
          Dim Index As Long, var As Long
53        Index = 0
          
          ' We loop through the variables and arrange them in the required order:
          '     1st - non-linear continuous
          '     2nd - non-linear integer
          '     3rd - linear arcs (N/A)
          '     4th - other linear
          '     5th - binary
          '     6th - other integer
          
          ' Non-linear continuous
54        For i = 1 To NonLinearContinuous.Count
55            UpdateStatusBar "OpenSolver: Creating .nl file. Outputting non-linear continuous vars"
              
56            var = NonLinearContinuous(i)
57            AddVariable CellNames(var), Index, var
58        Next i
          
          ' Non-linear integer
59        For i = 1 To NonLinearInteger.Count
60            UpdateStatusBar "OpenSolver: Creating .nl file. Outputting non-linear integer vars"
              
61            var = NonLinearInteger(i)
62            AddVariable CellNames(var), Index, var
63        Next i
          
          ' Linear continuous
64        For i = 1 To LinearContinuous.Count
65            UpdateStatusBar "OpenSolver: Creating .nl file. Outputting linear continuous vars"

66            var = LinearContinuous(i)
67            AddVariable CellNames(var), Index, var
68        Next i
          
          ' Linear binary
69        For i = 1 To LinearBinary.Count
70            UpdateStatusBar "OpenSolver: Creating .nl file. Outputting linear binary vars"
              
71            var = LinearBinary(i)
72            AddVariable CellNames(var), Index, var
73        Next i
          
          ' Linear integer
74        For i = 1 To LinearInteger.Count
75            UpdateStatusBar "OpenSolver: Creating .nl file. Outputting linear integer vars"

76            var = LinearInteger(i)
77            AddVariable CellNames(var), Index, var
78        Next i
          
          ' ==============================================
          ' Update model stats
79        If Not SolveRelaxation Then
80            nbv = LinearBinary.Count
81            niv = LinearInteger.Count
82            nlvci = NonLinearInteger.Count
83        End If

ExitSub:
84        If RaiseError Then RethrowError
85        Exit Sub

ErrorHandler:
86        If Not ReportError("SolverFileNL", "MakeVariableMap") Then Resume
87        RaiseError = True
88        GoTo ExitSub
End Sub

' Adds a variable to all the variable maps with:
'   variable name:            cellName
'   .nl variable index:       index
'   parsed variable index:    i
Private Sub AddVariable(cellName As String, Index As Long, i As Long)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          ' Update variable maps
3         VariableMap.Add Key:=cellName, Item:=Index
4         VariableMapRev(Index) = cellName
5         VariableNLIndexToCollectionIndex(Index) = i
6         VariableCollectionIndexToNLIndex(i) = Index
          
          ' Update max length of variable name
7         If WriteComments Then
8             If Len(cellName) > maxcolnamelen_ Then
9                 maxcolnamelen_ = Len(cellName)
10            End If
11        End If
          
          ' Increase index for the next variable
12        Index = Index + 1

ExitSub:
13        If RaiseError Then RethrowError
14        Exit Sub

ErrorHandler:
15        If Not ReportError("SolverFileNL", "AddVariable") Then Resume
16        RaiseError = True
17        GoTo ExitSub
End Sub

' Creates maps from constraint name (e.g. c1_Test1_D4) to .nl constraint index (0 to n_con - 1) and vice-versa, and
' map from .nl constraint index to parsed constraint index
Private Sub MakeConstraintMap()
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         ReDim ConstraintMapRev(n_con) As String
4         ReDim ConstraintIndexToTreeIndex(n_con) As Long
          
          Dim Index As Long, i As Long, cellName As String
5         Index = 0
          
          ' We loop through the constraints and arrange them in the required order:
          '     1st - non-linear
          '     2nd - non-linear network (N/A)
          '     3rd - linear network (N/A)
          '     4th - linear
          
          ' Non-linear constraints
6         For i = 1 To n_con
7             UpdateStatusBar "OpenSolver: Creating .nl file. Outputting non-linear constraints " & i & "/" & n_con

8             If NonLinearConstraints(i) Then
                  ' Actual constraints
9                 If i <= numActualCons Then
10                    cellName = "c" & i & "_" & m.LHSKeys(i)
                  ' Formulae constraints
11                ElseIf i <= numActualCons + numFakeCons Then
12                    cellName = "f" & i & "_" & m.Formulae(i - numActualCons).strAddress
13                Else
14                    cellName = "seek_obj_" & ConvertCellToStandardName(s.ObjRange)
15                End If
16                AddConstraint cellName, Index, i
17            End If
18        Next i
          
          ' Linear constraints
19        For i = 1 To n_con
20            UpdateStatusBar "OpenSolver: Creating .nl file. Outputting linear constraints " & i & "/" & n_con

21            If Not NonLinearConstraints(i) Then
                  ' Actual constraints
22                If i <= numActualCons Then
23                    cellName = "c" & i & "_" & m.LHSKeys(i)
                  ' Formulae constraints
24                ElseIf i <= numActualCons + numFakeCons Then
25                    cellName = "f" & i & "_" & m.Formulae(i - numActualCons).strAddress
26                Else
27                    cellName = "seek_obj_" & ConvertCellToStandardName(s.ObjRange)
28                End If
29                AddConstraint cellName, Index, i
30            End If
31        Next i

ExitSub:
32        If RaiseError Then RethrowError
33        Exit Sub

ErrorHandler:
34        If Not ReportError("SolverFileNL", "MakeConstraintMap") Then Resume
35        RaiseError = True
36        GoTo ExitSub
End Sub

' Adds a constraint to all the constraint maps with:
'   constraint name:          cellName
'   .nl constraint index:     index
'   parsed constraint index:  i
Private Sub AddConstraint(cellName As String, Index As Long, i As Long)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          ' Update constraint map
3         ConstraintMapRev(Index) = cellName
4         ConstraintIndexToTreeIndex(Index) = i
          
          ' Update max length
5         If WriteComments Then
6             If Len(cellName) > maxrownamelen_ Then
7                 maxrownamelen_ = Len(cellName)
8             End If
9         End If
          
          ' Increase index for next constraint
10        Index = Index + 1

ExitSub:
11        If RaiseError Then RethrowError
12        Exit Sub

ErrorHandler:
13        If Not ReportError("SolverFileNL", "AddConstraint") Then Resume
14        RaiseError = True
15        GoTo ExitSub
End Sub

' Processes all constraint formulae into the .nl model formats
Private Sub ProcessFormulae()
          Dim RaiseError As Boolean
          Dim UserMessage As String
          Dim StackTraceMessage As String
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         Erase NonLinearConstraintTrees
4         Erase LinearConstraints
5         ReDim NonLinearConstraintTrees(1 To n_con) As ExpressionTree
6         ReDim LinearConstraints(1 To n_con) As Dictionary
7         ReDim LinearConstants(1 To n_con) As Double
8         ReDim ConstraintRelations(1 To n_con) As RelationConsts
          
9         ReDim NonLinearVars(1 To n_var) As Boolean
10        ReDim NonLinearConstraints(1 To n_con) As Boolean
11        ReDim NonZeroConstraintCount(1 To n_var) As Long
          
          ' Loop through all constraints and process each
          Dim i As Long
12        On Error GoTo BadActualCon
13        For i = 1 To numActualCons
14            UpdateStatusBar "OpenSolver: Processing formulae into expression trees... " & i & "/" & n_con & " formulae."
15            ProcessSingleFormula m.RHSKeys(i), m.LHSKeys(i), m.RELs(i), i
16        Next i
          
17        On Error GoTo BadFakeCon
18        For i = 1 To numFakeCons
19            UpdateStatusBar "OpenSolver: Processing formulae into expression trees... " & i + numActualCons & "/" & n_con & " formulae."
20            ProcessSingleFormula m.Formulae(i).strFormulaParsed, m.Formulae(i).strAddress, RelationConsts.RelationEQ, i + numActualCons
21        Next i
          
ExitSub:
22        Application.StatusBar = False
23        If RaiseError Then RethrowError
24        Exit Sub

ErrorHandler:
25        If Not ReportError("SolverFileNL", "ProcessFormulae", UserMessage:=UserMessage, StackTraceMessage:=StackTraceMessage) Then Resume
26        RaiseError = True
27        GoTo ExitSub

BadActualCon:
28        UserMessage = "Non-linear parser failed while processing constraint " & m.RHSKeys(i) & RelationEnumToString(m.RELs(i)) & m.LHSKeys(i) & "."
29        StackTraceMessage = UserMessage
30        GoTo ErrorHandler
          
BadFakeCon:
31        UserMessage = "Non-linear parser failed while processing cell " & m.Formulae(i).strAddress & "."
32        StackTraceMessage = UserMessage
33        GoTo ErrorHandler
End Sub

' Processes a single constraint into .nl format. We require:
'     - a non-linear ExpressionTree for all non-linear parts of the equation
'     - a linear Dictionary for the linear parts of the equation
'     - a constant Double for the constant part of the equation
' We also use the results of processing to update some of the model statistics
Private Sub ProcessSingleFormula(RHSExpression As String, LHSVariable As String, Relation As RelationConsts, i As Long)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          ' Convert the string formula into an ExpressionTree object
          Dim Tree As ExpressionTree
3         Set Tree = ConvertFormulaToExpressionTree(RHSExpression)
          
          ' The .nl file needs a linear coefficient for every variable in the constraint - non-linear or otherwise
          ' We need a list of all variables in this constraint so that we can know to include them in the linear part of the constraint.
          Dim constraint As New Dictionary
4         Tree.ExtractVariables constraint

          Dim LinearTrees As New Collection
5         Tree.MarkLinearity
          
          ' Constants in .nl expression trees must be as simple as possible.
          ' We cannot have constant * constant in the tree, it must be replaced with a single constant
          ' We need to evalulate and pull up all constants that we can.
6         Tree.PullUpConstants
          
          ' Remove linear terms from non-linear trees
7         Tree.PruneLinearTrees LinearTrees
          
          ' Process linear trees to separate constants and variables
          Dim constant As Double
8         constant = 0
          Dim j As Long
9         For j = 1 To LinearTrees.Count
10            LinearTrees(j).ConvertLinearTreeToConstraint constraint, constant
11        Next j
          
          ' Check that our variable LHS exists
12        If Not VariableIndex.Exists(LHSVariable) Then
              ' We must have a constant formula as the LHS, we can evaluate and merge with the constant
              ' We are bringing the constant to the RHS so must subtract
13            If Not Left(LHSVariable, 1) = "=" Then LHSVariable = "=" & LHSVariable
14            constant = constant - s.sheet.Evaluate(LHSVariable)
15        Else
              ' Our constraint has a single term on the LHS and a formulae on the right.
              ' The single LHS term needs to be included in the expression as a linear term with coefficient 1
              ' Move variable from the LHS to the linear constraint with coefficient -1
              Dim VarKey As Long
16            VarKey = VariableIndex.Item(LHSVariable)
17            If constraint.Exists(VarKey) Then
18                constraint.Item(VarKey) = constraint.Item(VarKey) - 1
19            Else
20                constraint.Add Key:=VarKey, Item:=-1
21            End If
22        End If
          
          ' Keep constant on RHS and take everything else to LHS.
          ' Need to flip all sign of elements in this constraint

          ' Flip all coefficients in linear constraint
23        InvertCoefficients constraint

          ' Negate non-linear tree
24        Set Tree = Tree.Negate
          
          ' Save results of processing
25        Set NonLinearConstraintTrees(i) = Tree
26        Set LinearConstraints(i) = constraint
27        LinearConstants(i) = constant
28        ConstraintRelations(i) = Relation
          
          ' Mark any non-linear variables that we haven't seen before by extracting any remaining variables from the non-linear tree.
          ' Any variable still present in the constraint must be part of the non-linear section
          Dim TempConstraint As New Dictionary, var As Variant
29        Tree.ExtractVariables TempConstraint
30        For Each var In TempConstraint.Keys
31            If Not NonLinearVars(var) Then
32                NonLinearVars(var) = True
33                nlvc = nlvc + 1
34            End If
35        Next var
          
          ' Remove any zero coefficients from the linear constraint if the variable is not in the non-linear tree
36        For Each var In constraint.Keys
37            If Not NonLinearVars(var) And constraint.Item(var) = 0 Then
38                constraint.Remove var
39            End If
40        Next var
          
          ' Increase count of non-linear constraints if the non-linear tree is non-empty
          ' An empty tree has a single "0" node
41        If Tree.NodeText <> "0" Then
42            NonLinearConstraints(i) = True
43            nlc = nlc + 1
44        End If
          
          ' Update jacobian counts using the linear variables present
45        For Each var In constraint.Keys
              ' The nl documentation says that the jacobian counts relate to "the numbers of nonzeros in the first n var - 1 columns of the Jacobian matrix"
              ' This means we should increases the count for any variable that is present and has a non-zero coefficient
              ' However, .nl files generated by AMPL seem to increase the count for any present variable, even if the coefficient is zero (i.e. non-linear variables)
              ' We adopt the AMPL behaviour as it gives faster solve times (see Test 40). Swap the If conditions below to change this behaviour
              
              'If constraint.Coefficient(j) <> 0 Then
              Dim VarIndex As Long
46            VarIndex = CLng(var)
47            NonZeroConstraintCount(VarIndex) = NonZeroConstraintCount(VarIndex) + 1
48            nzc = nzc + 1
              'End If
49        Next var

ExitSub:
50        If RaiseError Then RethrowError
51        Exit Sub

ErrorHandler:
52        If Not ReportError("SolverFileNL", "ProcessSingleFormula") Then Resume
53        RaiseError = True
54        GoTo ExitSub
End Sub

Sub InvertCoefficients(ByRef constraint As Dictionary)
          Dim Key As Variant
1         For Each Key In constraint.Keys
2             constraint.Item(Key) = -constraint.Item(Key)
3         Next Key
End Sub

' Process objective function into .nl format. We require the same as for constraints:
'     - a non-linear ExpressionTree for all non-linear parts of the equation
'     - a linear Dictionary for the linear parts of the equation
'     - a constant Double for the constant part of the equation
Private Sub ProcessObjective()
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
3         If s.ObjectiveSense = TargetObjective Then
              ' Instead of adding an objective, we add a constraint
4             ProcessSingleFormula s.ObjectiveTargetValue, ConvertCellToStandardName(s.ObjRange), RelationEQ, n_con
5         ElseIf s.ObjRange Is Nothing Then
              ' Do nothing is objective is missing
6         Else
              ' =======================================================
              ' Currently just one objective - a single linear variable
              ' We could move to multiple objectives if OpenSolver supported this
              
              ' Adjust objective count
7             n_obj = n_obj + 1

8             Erase NonLinearObjectiveTrees
9             Erase LinearObjectives
10            ReDim NonLinearObjectiveTrees(1 To n_obj) As ExpressionTree
11            ReDim ObjectiveSenses(1 To n_obj) As ObjectiveSenseType
12            ReDim ObjectiveCells(1 To n_obj) As String
13            ReDim LinearObjectives(1 To n_obj) As Dictionary

14            ObjectiveCells(1) = ConvertCellToStandardName(s.ObjRange)
15            ObjectiveSenses(1) = s.ObjectiveSense
              
              ' Objective non-linear constraint tree is empty
16            Set NonLinearObjectiveTrees(1) = CreateTree("0", ExpressionTreeNodeType.ExpressionTreeNumber)
              
              ' Objective has a single linear term - the objective variable with coefficient 1
              Dim Objective As New Dictionary
17            Objective.Add Key:=VariableIndex.Item(ObjectiveCells(1)), Item:=1
              
              ' Save results
18            Set LinearObjectives(1) = Objective
              
              ' Track non-zero jacobian count in objective
19            nzo = nzo + 1
20        End If

ExitSub:
21        If RaiseError Then RethrowError
22        Exit Sub

ErrorHandler:
23        If Not ReportError("SolverFileNL", "ProcessObjective") Then Resume
24        RaiseError = True
25        GoTo ExitSub
End Sub

' Writes header block for .nl file. This contains the model statistics
Private Sub MakeHeader()
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
3         Print #1, "g3 1 1 0"; vbTab; "# problem " & problem_name
4         Print #1, " " & n_var & " " & n_con & " " & n_obj & " " & nranges & " " & n_eqn_ & IIf(n_lcon > 0, " " & n_lcon, vbNullString); vbTab; "# vars, constraints, objectives, ranges, eqns"
5         Print #1, " " & nlc & " " & nlo; vbTab; "# nonlinear constraints, objectives"
6         Print #1, " " & nlnc & " " & lnc; vbTab; "# network constraints: nonlinear, linear"
7         Print #1, " " & nlvc & " " & nlvo & " " & nlvb; vbTab; "# nonlinear vars in constraints, objectives, both"
8         Print #1, " " & nwv_ & " " & nfunc_ & " " & arith & " " & flags; vbTab; "# linear network variables; functions; arith, flags"
9         Print #1, " " & nbv & " " & niv & " " & nlvbi & " " & nlvci & " " & nlvoi; vbTab; "# discrete variables: binary, integer, nonlinear (b,c,o)"
10        Print #1, " " & nzc & " " & nzo; vbTab; "# nonzeros in Jacobian, gradients"
11        Print #1, " " & maxrownamelen_ & " " & maxcolnamelen_; vbTab; "# max name lengths: constraints, variables"
12        Print #1, " " & comb & " " & comc & " " & como & " " & comc1 & " " & como1; vbTab; "# common exprs: b,c,o,c1,o1"

ExitSub:
13        If RaiseError Then RethrowError
14        Exit Sub

ErrorHandler:
15        If Not ReportError("SolverFileNL", "MakeHeader") Then Resume
16        RaiseError = True
17        GoTo ExitSub
End Sub

' Writes C blocks for .nl file. These describe the non-linear parts of each constraint.
Private Sub MakeCBlocks()
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim i As Long
3         For i = 1 To n_con
4             UpdateStatusBar "OpenSolver: Creating .nl file. Writing non-linear constraints " & i & "/" & n_con

              ' Add block header for the constraint
5             OutputLine 1, _
                  "C" & i - 1, _
                  "CONSTRAINT NON-LINEAR SECTION " & ConstraintMapRev(i - 1)
              
              ' Add expression tree
              Dim Tree As ExpressionTree
6             Set Tree = NonLinearConstraintTrees(ConstraintIndexToTreeIndex(i - 1))
7             Tree.ConvertToNL 1, CommentIndent
8         Next i

ExitSub:
9         If RaiseError Then RethrowError
10        Exit Sub

ErrorHandler:
11        If Not ReportError("SolverFileNL", "MakeCBlocks") Then Resume
12        RaiseError = True
13        GoTo ExitSub
End Sub

' Writes O blocks for .nl file. These describe the non-linear parts of each objective.
Private Sub MakeOBlocks()
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim i As Long
3         For i = 1 To n_obj
              ' Add block header for the objective
4             OutputLine 1, _
                  "O" & i - 1 & " " & ConvertObjectiveSenseToNL(ObjectiveSenses(i)), _
                  "OBJECTIVE NON-LINEAR SECTION " & ObjectiveCells(i)
              
              ' Add expression tree
              Dim Tree As ExpressionTree
5             Set Tree = NonLinearObjectiveTrees(i)
6             Tree.ConvertToNL 1, CommentIndent
7         Next i

ExitSub:
8         If RaiseError Then RethrowError
9         Exit Sub

ErrorHandler:
10        If Not ReportError("SolverFileNL", "MakeOBlocks") Then Resume
11        RaiseError = True
12        GoTo ExitSub
End Sub

' Writes D block for .nl file. This contains the initial guess for dual variables.
' We don't use this, so just set them all to zero
Private Sub MakeDBlock()
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          ' Add block header
3         OutputLine 1, "d" & n_con, "INITIAL DUAL GUESS"
          
          ' Set duals to zero for all constraints
          Dim i As Long
4         For i = 1 To n_con
5             UpdateStatusBar "OpenSolver: Creating .nl file. Writing initial duals " & i & "/" & n_con
              
6             OutputLine 1, _
                  i - 1 & " 0", _
                  "    " & ConstraintMapRev(i - 1) & " = " & 0
7         Next i

ExitSub:
8         If RaiseError Then RethrowError
9         Exit Sub

ErrorHandler:
10        If Not ReportError("SolverFileNL", "MakeDBlock") Then Resume
11        RaiseError = True
12        GoTo ExitSub
End Sub

' Writes X block for .nl file. This contains the initial guess for primal variables
Private Sub MakeXBlock()
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          ' Add block header
3         OutputLine 1, "x" & n_var, "INITIAL PRIMAL GUESS"

          ' Loop through the variables in .nl variable order
          Dim i As Long, initial As Double, VarIndex As Long
4         For i = 1 To n_var
5             UpdateStatusBar "OpenSolver: Creating .nl file. Writing initial values " & i & "/" & n_var
              
6             VarIndex = VariableNLIndexToCollectionIndex(i - 1)
              
              ' Get initial values
7             If VarIndex <= numActualVars Then
                  ' Actual variables - use the value in the actual cell
8                 initial = InitialVariableValues(VarIndex)
9             Else
                  ' Formulae variables - use the initial value saved in the CFormula instance
10                initial = m.Formulae(VarIndex - numActualVars).initialValue
11            End If

12            OutputLine 1, _
                  i - 1 & " " & StrExNoPlus(initial), _
                  "    " & VariableMapRev(i - 1) & " = " & initial
13        Next i

ExitSub:
14        If RaiseError Then RethrowError
15        Exit Sub

ErrorHandler:
16        If Not ReportError("SolverFileNL", "MakeXBlock") Then Resume
17        RaiseError = True
18        GoTo ExitSub
End Sub

' Writes R block for .nl file. This contains the constant values for each constraint and the relation type
Private Sub MakeRBlock()
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If n_con = 0 Then GoTo ExitSub

           ' Add block header
4         OutputLine 1, "r", "CONSTRAINT BOUNDS"
          
          ' Apply bounds according to the relation type
          Dim i As Long, BoundType As Long, Comment As String, bound As Double
5         For i = 1 To n_con
6             UpdateStatusBar "OpenSolver: Creating .nl file. Writing constraint bounds " & i & "/" & n_con
              
7             bound = LinearConstants(ConstraintIndexToTreeIndex(i - 1))
8             ConvertConstraintToNL ConstraintRelations(ConstraintIndexToTreeIndex(i - 1)), BoundType, Comment

9             OutputLine 1, _
                  BoundType & " " & StrExNoPlus(bound), _
                  "    " & ConstraintMapRev(i - 1) & Comment & bound
10        Next i

ExitSub:
11        If RaiseError Then RethrowError
12        Exit Sub

ErrorHandler:
13        If Not ReportError("SolverFileNL", "MakeRBlock") Then Resume
14        RaiseError = True
15        GoTo ExitSub
End Sub

' Writes B block for .nl file. This contains the variable bounds
Private Sub MakeBBlock()
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          ' Write block header
3         OutputLine 1, "b", "VARIABLE BOUNDS"
          
          Dim i As Long, bound As String, Comment As String, VarIndex As Long, VarName As String
4         For i = 1 To n_var
5             UpdateStatusBar "OpenSolver: Creating .nl file. Writing variable bounds " & i & "/" & n_var
              
6             VarIndex = VariableNLIndexToCollectionIndex(i - 1)
7             Comment = "    " & VariableMapRev(i - 1)
           
8             If VarIndex <= numActualVars Then
9                 If BinaryVars(VarIndex) Then
10                  bound = "0 0 1"
11                  Comment = Comment & " IN [0, 1]"
                  ' Real variables, use actual bounds
12                ElseIf s.AssumeNonNegativeVars Then
13                    VarName = s.VarName(VarIndex)
14                    If s.VarLowerBounds.Exists(VarName) And Not BinaryVars(VarIndex) Then
15                        bound = "3"
16                        Comment = Comment & " FREE"
17                    Else
18                        bound = "2 0"
19                        Comment = Comment & " >= 0"
20                    End If
21                Else
22                    bound = "3"
23                    Comment = Comment & " FREE"
24                End If
25            Else
                  ' Fake formulae variables - no bounds
26                bound = "3"
27                Comment = Comment & " FREE"
28            End If
29            OutputLine 1, bound, Comment
30        Next i

ExitSub:
31        If RaiseError Then RethrowError
32        Exit Sub

ErrorHandler:
33        If Not ReportError("SolverFileNL", "MakeBBlock") Then Resume
34        RaiseError = True
35        GoTo ExitSub
End Sub

' Writes K block for .nl file. This contains the cumulative count of non-zero jacobian entries for the first n-1 variables
Private Sub MakeKBlock()
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If n_var = 0 Then GoTo ExitSub

          ' Add block header
4         OutputLine 1, _
              "k" & n_var - 1, _
              "NUMBER OF JACOBIAN ENTRIES (CUMULATIVE) FOR FIRST " & n_var - 1 & " VARIABLES"
          
          ' Loop through first n_var - 1 variables and add the non-zero count to the running total
          Dim i As Long, total As Long
5         total = 0
6         For i = 1 To n_var - 1
7             UpdateStatusBar "OpenSolver: Creating .nl file. Writing jacobian counts " & i & "/" & n_var - 1
              
8             total = total + NonZeroConstraintCount(VariableNLIndexToCollectionIndex(i - 1))

9             OutputLine 1, _
                  CStr(total), _
                  "    Up to " & VariableMapRev(i - 1) & ": " & total & " entries in Jacobian"
10        Next i

ExitSub:
11        If RaiseError Then RethrowError
12        Exit Sub

ErrorHandler:
13        If Not ReportError("SolverFileNL", "MakeKBlock") Then Resume
14        RaiseError = True
15        GoTo ExitSub
End Sub

' Writes J blocks for .nl file. These contain the linear part of each constraint
Private Sub MakeJBlocks()
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim i As Long, TreeIndex As Long, VarIndex As Long
3         For i = 1 To n_con
4             UpdateStatusBar "OpenSolver: Creating .nl file. Writing linear constraints " & i & "/" & n_con
              
5             TreeIndex = ConstraintIndexToTreeIndex(i - 1)
              Dim LinearConstraint As Dictionary
6             Set LinearConstraint = LinearConstraints(TreeIndex)
          
              ' Make header
7             OutputLine 1, _
                  "J" & i - 1 & " " & LinearConstraint.Count, _
                  "CONSTRAINT LINEAR SECTION " & ConstraintMapRev(i - 1)
              
              ' We need variables output in .nl order
              Dim j As Long
8             For j = 0 To n_var - 1
9                 VarIndex = VariableNLIndexToCollectionIndex(j)
10                If LinearConstraint.Exists(VarIndex) Then
11                    OutputLine 1, _
                          CStr(j) & " " & StrExNoPlus(LinearConstraint.Item(VarIndex)), _
                          "    + " & LinearConstraint.Item(VarIndex) & " * " & VariableMapRev(j)
12                End If
13            Next j
14        Next i

ExitSub:
15        If RaiseError Then RethrowError
16        Exit Sub

ErrorHandler:
17        If Not ReportError("SolverFileNL", "MakeJBlocks") Then Resume
18        RaiseError = True
19        GoTo ExitSub
End Sub

' Writes the G blocks for .nl file. These contain the linear parts of each objective
Private Sub MakeGBlocks()
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim i As Long
3         For i = 1 To n_obj
              ' Make header
4             OutputLine 1, _
                  "G" & i - 1 & " " & LinearObjectives(i).Count, _
                  "OBJECTIVE LINEAR SECTION " & ObjectiveCells(i)
              
              Dim LinearObjective As Dictionary
5             Set LinearObjective = LinearObjectives(i)
              
              ' This loop is not in the right order (see J blocks)
              ' Since the objective only containts one variable, the output will still be correct without reordering
              Dim j As Long, VarIndex As Long
6             For j = 0 To n_var - 1
7                 VarIndex = VariableNLIndexToCollectionIndex(j)
8                 If LinearObjective.Exists(VarIndex) Then
9                     OutputLine 1, _
                          CStr(j) & " " & StrExNoPlus(LinearObjective.Item(VarIndex)), _
                          "    + " & LinearObjective.Item(VarIndex) & " * " & VariableMapRev(j)
10                End If
11            Next j
12        Next i

ExitSub:
13        If RaiseError Then RethrowError
14        Exit Sub

ErrorHandler:
15        If Not ReportError("SolverFileNL", "MakeGBlocks") Then Resume
16        RaiseError = True
17        GoTo ExitSub
End Sub

' Writes the .col summary file. This contains the variable names listed in .nl order
Private Sub OutputColFile()
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim ColFilePathName As String
3         GetTempFilePath "model.col", ColFilePathName
          
4         DeleteFileAndVerify ColFilePathName

5         Open ColFilePathName For Output As #2
          
6         UpdateStatusBar "OpenSolver: Creating .nl file. Writing col file"
          Dim i As Long
7         For i = 1 To n_var
8             WriteToFile 2, VariableMapRev(i)
9         Next i
          
ExitSub:
10        Close #2
11        If RaiseError Then RethrowError
12        Exit Sub

ErrorHandler:
13        If Not ReportError("SolverFileNL", "OutputColFile") Then Resume
14        RaiseError = True
15        GoTo ExitSub
End Sub

' Writes the .row summary file. This contains the constraint names listed in .nl order
Private Sub OutputRowFile()
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim RowFilePathName As String
3         GetTempFilePath "model.row", RowFilePathName
          
4         DeleteFileAndVerify RowFilePathName

5         Open RowFilePathName For Output As #3
          
6         UpdateStatusBar "OpenSolver: Creating .nl file. Writing con file"
          Dim i As Long
7         For i = 1 To n_con
8             DoEvents
9             WriteToFile 3, ConstraintMapRev(i)
10        Next i

ExitSub:
11        Close #3
12        If RaiseError Then RethrowError
13        Exit Sub

ErrorHandler:
14        If Not ReportError("SolverFileNL", "OutputRowFile") Then Resume
15        RaiseError = True
16        GoTo ExitSub
End Sub


' Adds a new line to the current string, appending LineText at position 0 and CommentText at position CommentSpacing
Sub OutputLine(FileNum As Long, LineText As String, Optional CommentText As String = vbNullString)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         Print #FileNum, LineText;
          
          ' Add comment with padding if comment should be included
4         If WriteComments And Len(CommentText) > 0 Then
5             Print #FileNum, Tab(CommentSpacing); "# " & CommentText;
6         End If
          
7         Print #FileNum,

ExitSub:
8         If RaiseError Then RethrowError
9         Exit Sub

ErrorHandler:
10        If Not ReportError("SolverFileNL", "AddNewLine") Then Resume
11        RaiseError = True
12        GoTo ExitSub
End Sub

Private Function ConvertFormulaToExpressionTree(strFormula As String) As ExpressionTree
          ' Converts a string formula to a complete expression tree
          ' Uses the Shunting Yard algorithm (adapted to produce an expression tree) which takes O(n) time
          ' https://en.wikipedia.org/wiki/Shunting-yard_algorithm
          ' For details on modifications to algorithm
          ' http://wcipeg.com/wiki/Shunting_yard_algorithm#Conversion_into_syntax_tree

          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim tksFormula As Tokens
3         Set tksFormula = ParseFormula("=" + strFormula)
          
          Dim Operands As New ExpressionTreeStack, Operators As New StringStack, ArgCounts As New OperatorArgCountStack
          
          Dim i As Long, tkn As Token, tknOld As String, Tree As ExpressionTree
4         For i = 1 To tksFormula.Count
5             Set tkn = tksFormula.Item(i)
              
6             Select Case tkn.TokenType
              ' If the token is a number or variable, then add it to the operands stack as a new tree.
              Case TokenType.Number
                  ' Might be a negative number, if so we need to parse out the neg operator
7                 Set Tree = CreateTree(tkn.Text, ExpressionTreeNumber)
8                 If Left(Tree.NodeText, 1) = "-" Then
9                     AddNegToTree Tree
10                End If
11                Operands.Push Tree
                      
12            Case TokenType.Text
13                Operands.Push CreateTree(tkn.Text, ExpressionTreeString)
14            Case TokenType.Reference
                  ' TODO this is a hacky way of distinguishing strings eg "A1" from model variables "Sheet_A1"
                  ' Obviously it will fail if the string has an underscore.
15                If InStr(tkn.Text, "_") > 0 Then
16                    Operands.Push CreateTree(tkn.Text, ExpressionTreeVariable)
17                Else
18                    Operands.Push CreateTree(tkn.Text, ExpressionTreeString)
19                End If
                  
              ' If the token is a function token, then push it onto the operators stack along with a left parenthesis (tokeniser strips the parenthesis).
20            Case TokenType.FunctionOpen
21                Operators.Push ConvertExcelFunctionToNL(tkn.Text)
22                Operators.Push "("
                  
                  ' Start a new argument count
23                ArgCounts.PushNewCount
                  
              ' If the token is a function argument separator (e.g., a comma)
24            Case TokenType.ParameterSeparator
                  ' Until the token at the top of the operator stack is a left parenthesis, pop operators off the stack onto the operands stack as a new tree.
                  ' If no left parentheses are encountered, either the separator was misplaced or parentheses were mismatched.
25                Do While Operators.Peek() <> "("
26                    PopOperator Operators, Operands
                      
                      ' If the operator stack runs out without finding a left parenthesis, then there are mismatched parentheses
27                    If Operators.Count = 0 Then
28                        RaiseGeneralError "Mismatched parentheses"
29                    End If
30                Loop
                  ' Increase arg count for the new parameter
31                ArgCounts.Increase
              
              ' If the token is an operator
32            Case TokenType.ArithmeticOperator, TokenType.UnaryOperator, TokenType.ComparisonOperator
33                If tkn.TokenType = TokenType.UnaryOperator Then
34                    Select Case tkn.Text
                      Case "-"
                          ' Mark as unary minus
35                        tkn.Text = "neg"
36                    Case "+"
                          ' Discard unary plus and move to next element
37                        GoTo NextToken
38                    Case Else
                          ' Unknown unary operator
39                        RaiseGeneralError "While parsing formula for .nl output, the following unary operator was encountered: " & tkn.Text & vbNewLine & vbNewLine & _
                                            "The entire formula was: " & vbNewLine & _
                                            "=" & strFormula
40                    End Select
41                Else
42                    tkn.Text = ConvertExcelFunctionToNL(tkn.Text)
43                End If
                  ' While there is an operator token at the top of the operator stack
44                Do While Operators.Count > 0
45                    tknOld = Operators.Peek()
                      ' If either tkn is left-associative and its precedence is less than or equal to that of tknOld
                      ' or tkn has precedence less than that of tknOld
46                    If CheckPrecedence(tkn.Text, tknOld) Then
                          ' Pop tknOld off the operator stack onto the operand stack as a new tree
47                        PopOperator Operators, Operands
48                    Else
49                        Exit Do
50                    End If
51                Loop
                  ' Push operator onto the operator stack
52                Operators.Push tkn.Text
                  
              ' If the token is a left parenthesis, then push it onto the operator stack
53            Case TokenType.SubExpressionOpen
54                Operators.Push tkn.Text
                  
              ' If the token is a right parenthesis
55            Case TokenType.SubExpressionClose, TokenType.FunctionClose
                  ' Until the token at the top of the operator stack is not a left parenthesis, pop operators off the stack onto the operand stack as a new tree.
56                Do While Operators.Peek <> "("
57                    PopOperator Operators, Operands
                      ' If the operator stack runs out without finding a left parenthesis, then there are mismatched parentheses
58                    If Operators.Count = 0 Then
59                        RaiseGeneralError "Mismatched parentheses"
60                    End If
61                Loop
                  ' Pop the left parenthesis from the operator stack, but not onto the operand stack
62                Operators.Pop
                  ' If the token at the top of the stack is a function token, pop it onto the operand stack as a new tree
63                If Operators.Count > 0 Then
64                    If IsFunctionOperator(Operators.Peek()) Then
65                        PopOperator Operators, Operands, ArgCounts.PopCount
66                    End If
67                End If
68            End Select
NextToken:
69        Next i
          
          ' While there are still tokens in the operator stack
70        Do While Operators.Count > 0
              ' If the token on the top of the operator stack is a parenthesis, then there are mismatched parentheses
71            If Operators.Peek = "(" Then
72                RaiseGeneralError "Mismatched parentheses"
73            End If
              ' Pop the operator onto the operand stack as a new tree
74            PopOperator Operators, Operands
75        Loop
          
          ' We are left with a single tree in the operand stack - this is the complete expression tree
76        Set ConvertFormulaToExpressionTree = Operands.Pop

ExitFunction:
77        If RaiseError Then RethrowError
78        Exit Function

ErrorHandler:
79        If Not ReportError("SolverFileNL", "ConvertFormulaToExpressionTree") Then Resume
80        RaiseError = True
81        GoTo ExitFunction
          
End Function

' Creates an ExpressionTree and initialises NodeText and NodeType
Public Function CreateTree(NodeText As String, NodeType As Long) As ExpressionTree
          Dim obj As ExpressionTree

          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         Set obj = New ExpressionTree
4         obj.NodeText = NodeText
5         obj.NodeType = NodeType

6         Set CreateTree = obj

ExitFunction:
7         If RaiseError Then RethrowError
8         Exit Function

ErrorHandler:
9         If Not ReportError("SolverFileNL", "CreateTree") Then Resume
10        RaiseError = True
11        GoTo ExitFunction
End Function

Function IsNAry(FunctionName As String) As Boolean
1         Select Case FunctionName
          Case "min", "max", "sum", "count", "numberof", "numberofs", "and_n", "or_n", "alldiff"
2             IsNAry = True
3         Case Else
4             IsNAry = False
5         End Select
End Function

' Determines the number of operands expected by a .nl operator
Private Function NumberOfOperands(FunctionName As String, Optional ArgCount As Long = 0) As Long
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         Select Case FunctionName
          Case "floor", "ceil", "abs", "neg", "not", "tanh", "tan", "sqrt", "sinh", "sin", "log10", "log", "exp", "cosh", "cos", "atanh", "atan", "asinh", "asin", "acosh", "acos"
4             NumberOfOperands = 1
5         Case "plus", "minus", "mult", "div", "rem", "pow", "less", "or", "and", "lt", "le", "eq", "ge", "gt", "ne", "atan2", "intdiv", "precision", "round", "trunc", "iff"
6             NumberOfOperands = 2
7         Case "min", "max", "sum", "count", "numberof", "numberofs", "and_n", "or_n", "alldiff"
              'n-ary operator, read number of args from the arg counter
8             NumberOfOperands = ArgCount
9         Case "if", "ifs", "implies"
10            NumberOfOperands = 3
11        Case Else
12            RaiseGeneralError "Unknown function " & FunctionName & vbCrLf & "Please let us know about this so we can fix it."
13        End Select

ExitFunction:
14        If RaiseError Then RethrowError
15        Exit Function

ErrorHandler:
16        If Not ReportError("SolverFileNL", "NumberOfOperands") Then Resume
17        RaiseError = True
18        GoTo ExitFunction
End Function

' Converts common Excel functions to .nl operators
Private Function ConvertExcelFunctionToNL(FunctionName As String) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         FunctionName = LCase(FunctionName)
4         Select Case FunctionName
          Case "ln":       FunctionName = "log"
5         Case "+":        FunctionName = "plus"
6         Case "-":        FunctionName = "minus"
7         Case "*":        FunctionName = "mult"
8         Case "/":        FunctionName = "div"
9         Case "mod":      FunctionName = "rem"
10        Case "^":        FunctionName = "pow"
11        Case "<":        FunctionName = "lt"
12        Case "<=":       FunctionName = "le"
13        Case "=":        FunctionName = "eq"
14        Case ">=":       FunctionName = "ge"
15        Case ">":        FunctionName = "gt"
16        Case "<>":       FunctionName = "ne"
17        Case "quotient": FunctionName = "intdiv"
18        Case "and":      FunctionName = "and_n"
19        Case "or":       FunctionName = "or_n"
              
20        Case "log", "ceiling", "floor", "power"
21            RaiseGeneralError "Not implemented yet: " & FunctionName & vbCrLf & "Please let us know about this so we can fix it."
22        End Select
23        ConvertExcelFunctionToNL = FunctionName

ExitFunction:
24        If RaiseError Then RethrowError
25        Exit Function

ErrorHandler:
26        If Not ReportError("SolverFileNL", "ConvertExcelFunctionToNL") Then Resume
27        RaiseError = True
28        GoTo ExitFunction
End Function

' Determines the precedence of arithmetic operators
Private Function Precedence(tkn As String) As Long
1         Select Case tkn
          Case "eq", "ne", "gt", "ge", "lt", "le"
2             Precedence = 1
3         Case "plus", "minus", "neg"
4             Precedence = 2
5         Case "mult", "div"
6             Precedence = 3
7         Case "pow"
8             Precedence = 4
9         Case "neg"
10            Precedence = 5
11        Case Else
12            Precedence = -1
13        End Select
End Function

' Checks the precedence of two operators to determine if the current operator on the stack should be popped
Private Function CheckPrecedence(tkn1 As String, tkn2 As String) As Boolean
          ' Either tkn1 is left-associative and its precedence is less than or equal to that of tkn2
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If OperatorIsLeftAssociative(tkn1) Then
4             If Precedence(tkn1) <= Precedence(tkn2) Then
5                 CheckPrecedence = True
6             Else
7                 CheckPrecedence = False
8             End If
          ' Or tkn1 has precedence less than that of tkn2
9         Else
10            If Precedence(tkn1) < Precedence(tkn2) Then
11                CheckPrecedence = True
12            Else
13                CheckPrecedence = False
14            End If
15        End If

ExitFunction:
16        If RaiseError Then RethrowError
17        Exit Function

ErrorHandler:
18        If Not ReportError("SolverFileNL", "CheckPrecedence") Then Resume
19        RaiseError = True
20        GoTo ExitFunction
End Function

' Determines the left-associativity of arithmetic operators
Private Function OperatorIsLeftAssociative(tkn As String) As Boolean
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         Select Case tkn
          Case "plus", "minus", "mult", "div", "eq", "ne", "gt", "ge", "lt", "le", "neg"
4             OperatorIsLeftAssociative = True
5         Case "pow"
6             OperatorIsLeftAssociative = False
7         Case Else
8             RaiseGeneralError "Unknown associativity: " & tkn & vbNewLine & "Please let us know about this so we can fix it."
9         End Select

ExitFunction:
10        If RaiseError Then RethrowError
11        Exit Function

ErrorHandler:
12        If Not ReportError("SolverFileNL", "OperatorIsLeftAssociative") Then Resume
13        RaiseError = True
14        GoTo ExitFunction
End Function

' Pops an operator from the operator stack along with the corresponding number of operands.
Private Sub PopOperator(Operators As StringStack, Operands As ExpressionTreeStack, Optional ArgCount As Long = 0)
          ' Pop the operator and create a new ExpressionTree
          Dim Operator As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         Operator = Operators.Pop()
          Dim NewTree As ExpressionTree
4         Set NewTree = CreateTree(Operator, ExpressionTreeOperator)
          
          ' Pop the required number of operands from the operand stack and set as children of the new operator tree
          Dim NumToPop As Long, i As Long, Tree As ExpressionTree
5         NumToPop = NumberOfOperands(Operator, ArgCount)
6         For i = NumToPop To 1 Step -1
7             Set Tree = Operands.Pop
8             NewTree.SetChild i, Tree
9         Next i
          
          ' Add the new tree to the operands stack
10        Operands.Push NewTree

ExitSub:
11        If RaiseError Then RethrowError
12        Exit Sub

ErrorHandler:
13        If Not ReportError("SolverFileNL", "PopOperator") Then Resume
14        RaiseError = True
15        GoTo ExitSub
End Sub

' Check whether a token on the operator stack is a function operator (vs. an arithmetic operator)
Private Function IsFunctionOperator(tkn As String) As Boolean
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         Select Case tkn
          Case "plus", "minus", "mult", "div", "pow", "neg", "("
4             IsFunctionOperator = False
5         Case Else
6             IsFunctionOperator = True
7         End Select

ExitFunction:
8         If RaiseError Then RethrowError
9         Exit Function

ErrorHandler:
10        If Not ReportError("SolverFileNL", "IsFunctionOperator") Then Resume
11        RaiseError = True
12        GoTo ExitFunction
End Function

' Negates a tree by adding a 'neg' node to the root
Private Sub AddNegToTree(Tree As ExpressionTree)
          Dim NewTree As ExpressionTree
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         Set NewTree = CreateTree("neg", ExpressionTreeOperator)
          
4         Tree.NodeText = Mid(Tree.NodeText, 2)
5         NewTree.SetChild 1, Tree
          
6         Set Tree = NewTree

ExitSub:
7         If RaiseError Then RethrowError
8         Exit Sub

ErrorHandler:
9         If Not ReportError("SolverFileNL", "AddNegToTree") Then Resume
10        RaiseError = True
11        GoTo ExitSub
End Sub

' Formats an expression tree node's text as .nl output
Function FormatNL(NodeText As String, NodeType As ExpressionTreeNodeType) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         Select Case NodeType
          Case ExpressionTreeVariable
4             FormatNL = "v" & VariableMap.Item(NodeText)
5         Case ExpressionTreeNumber
6             FormatNL = "n" & StrExNoPlus(Val(NodeText))
7         Case ExpressionTreeOperator
8             FormatNL = "o" & CStr(ConvertOperatorToNLCode(NodeText))
9         End Select

ExitFunction:
10        If RaiseError Then RethrowError
11        Exit Function

ErrorHandler:
12        If Not ReportError("SolverFileNL", "FormatNL") Then Resume
13        RaiseError = True
14        GoTo ExitFunction
End Function

' Converts an operator string to .nl code
Private Function ConvertOperatorToNLCode(FunctionName As String) As Long
1         Select Case FunctionName
          Case "plus":      ConvertOperatorToNLCode = 0
2         Case "minus":     ConvertOperatorToNLCode = 1
3         Case "mult":      ConvertOperatorToNLCode = 2
4         Case "div":       ConvertOperatorToNLCode = 3
5         Case "rem":       ConvertOperatorToNLCode = 4
6         Case "pow":       ConvertOperatorToNLCode = 5
7         Case "less":      ConvertOperatorToNLCode = 6
8         Case "min":       ConvertOperatorToNLCode = 11
9         Case "max":       ConvertOperatorToNLCode = 12
10        Case "floor":     ConvertOperatorToNLCode = 13
11        Case "ceil":      ConvertOperatorToNLCode = 14
12        Case "abs":       ConvertOperatorToNLCode = 15
13        Case "neg":       ConvertOperatorToNLCode = 16
14        Case "or":        ConvertOperatorToNLCode = 20
15        Case "and":       ConvertOperatorToNLCode = 21
16        Case "lt":        ConvertOperatorToNLCode = 22
17        Case "le":        ConvertOperatorToNLCode = 23
18        Case "eq":        ConvertOperatorToNLCode = 24
19        Case "ge":        ConvertOperatorToNLCode = 28
20        Case "gt":        ConvertOperatorToNLCode = 29
21        Case "ne":        ConvertOperatorToNLCode = 30
22        Case "if":        ConvertOperatorToNLCode = 35
23        Case "not":       ConvertOperatorToNLCode = 34
24        Case "tanh":      ConvertOperatorToNLCode = 37
25        Case "tan":       ConvertOperatorToNLCode = 38
26        Case "sqrt":      ConvertOperatorToNLCode = 39
27        Case "sinh":      ConvertOperatorToNLCode = 40
28        Case "sin":       ConvertOperatorToNLCode = 41
29        Case "log10":     ConvertOperatorToNLCode = 42
30        Case "log":       ConvertOperatorToNLCode = 43
31        Case "exp":       ConvertOperatorToNLCode = 44
32        Case "cosh":      ConvertOperatorToNLCode = 45
33        Case "cos":       ConvertOperatorToNLCode = 46
34        Case "atanh":     ConvertOperatorToNLCode = 47
35        Case "atan2":     ConvertOperatorToNLCode = 48
36        Case "atan":      ConvertOperatorToNLCode = 49
37        Case "asinh":     ConvertOperatorToNLCode = 50
38        Case "asin":      ConvertOperatorToNLCode = 51
39        Case "acosh":     ConvertOperatorToNLCode = 52
40        Case "acos":      ConvertOperatorToNLCode = 53
41        Case "sum":       ConvertOperatorToNLCode = 54
42        Case "intdiv":    ConvertOperatorToNLCode = 55
43        Case "precision": ConvertOperatorToNLCode = 56
44        Case "round":     ConvertOperatorToNLCode = 57
45        Case "trunc":     ConvertOperatorToNLCode = 58
46        Case "count":     ConvertOperatorToNLCode = 59
47        Case "numberof":  ConvertOperatorToNLCode = 60
48        Case "numberofs": ConvertOperatorToNLCode = 61
49        Case "ifs":       ConvertOperatorToNLCode = 65
50        Case "and_n":     ConvertOperatorToNLCode = 70
51        Case "or_n":      ConvertOperatorToNLCode = 71
52        Case "implies":   ConvertOperatorToNLCode = 72
53        Case "iff":       ConvertOperatorToNLCode = 73
54        Case "alldiff":   ConvertOperatorToNLCode = 74
55        End Select
End Function

' Converts an objective sense to .nl code
Private Function ConvertObjectiveSenseToNL(ObjectiveSense As ObjectiveSenseType) As Long
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         Select Case ObjectiveSense
          Case ObjectiveSenseType.MaximiseObjective
4             ConvertObjectiveSenseToNL = 1
5         Case ObjectiveSenseType.MinimiseObjective
6             ConvertObjectiveSenseToNL = 0
7         Case Else
8             RaiseGeneralError "Objective sense not supported: " & ObjectiveSense
9         End Select

ExitFunction:
10        If RaiseError Then RethrowError
11        Exit Function

ErrorHandler:
12        If Not ReportError("SolverFileNL", "ConvertObjectiveSenseToNL") Then Resume
13        RaiseError = True
14        GoTo ExitFunction
End Function

' Converts RelationConsts enum to .nl code.
Private Sub ConvertConstraintToNL(Relation As RelationConsts, BoundType As Long, Comment As String)
1         Select Case Relation
              Case RelationConsts.RelationLE ' Upper Bound on LHS
2                 BoundType = 1
3                 Comment = " <= "
4             Case RelationConsts.RelationEQ ' Equality
5                 BoundType = 4
6                 Comment = " == "
7             Case RelationConsts.RelationGE ' Upper Bound on RHS
8                 BoundType = 2
9                 Comment = " >= "
10        End Select
End Sub

Sub ReadResults_NL(s As COpenSolver)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim solutionExpected As Boolean
3         s.SolutionWasLoaded = False
          
4         If Not FileOrDirExists(s.SolutionFilePathName) Then
5             If Not GetExtraInfoFromLog(s) Then
6                 CheckLog_NL s
7                 RaiseGeneralError _
                      "The solver did not create a solution file. No new solution is available." & vbNewLine & vbNewLine & _
                      "This can happen when the initial conditions are invalid. " & _
                      "Check the log file for more information.", _
                      LINK_NO_SOLUTION_FILE
8             End If
9         Else
10            UpdateStatusBar "OpenSolver: Reading Solution... ", True
          
              ' Reference implementation for reading .sol files:
              ' https://github.com/ampl/mp/blob/master/src/asl/solvers/readsol.c
              Dim Line As String
11            Open s.SolutionFilePathName For Input As #1
              
12            Do While True  ' Skip empty line at start of file
13                Line Input #1, Line
14                If Trim(Line) <> vbNullString Then Exit Do
15            Loop
              ' `Line` now has first line of solve message
              
              Dim SolveMessage As String
16            SolveMessage = vbNullString
17            Do While True
18                SolveMessage = SolveMessage & IIf(Len(SolveMessage) <> 0, vbNewLine, vbNullString) & Line
19                Line Input #1, Line
20                If Trim(Line) = vbNullString Then Exit Do
21            Loop
              
22            Line Input #1, Line
23            If LCase(Left(Line, 7)) <> "options" Then
24                RaiseGeneralError "Bad .sol file"
25            End If
              
              Dim Options(1 To 14) As Long
              Dim i As Long
26            For i = 1 To 3
27               Line Input #1, Line
28               Options(i) = Val(Line)
29            Next i
              
              Dim NumOptions As Long
30            NumOptions = Options(1)
31            If NumOptions < 3 Or NumOptions > 9 Then
32                RaiseGeneralError "Too many options"
33            End If
              Dim NeedVbTol As Boolean
34            NeedVbTol = False
35            If Options(3) = 3 Then
36                NumOptions = NumOptions - 2
37                NeedVbTol = True
38            End If
              
39            For i = 4 To NumOptions + 1
40                Line Input #1, Line
41                Options(i) = Val(Line)
42            Next i
              
43            Line Input #1, Line
44            If Val(Line) <> n_con Then
45                RaiseGeneralError "Wrong number of constraints"
46            End If
              
              Dim NumDualsToRead As Long
47            Line Input #1, Line
48            NumDualsToRead = Val(Line)
              
49            Line Input #1, Line
50            If Val(Line) <> n_var Then
51                RaiseGeneralError "Wrong number of variables"
52            End If
              
              Dim NumVarsToRead As Long
53            Line Input #1, Line
54            NumVarsToRead = Val(Line)
55            If NumVarsToRead > 0 Then s.SolutionWasLoaded = True
              
56            If NeedVbTol Then Line Input #1, Line  ' Throw away vbtol line if present
              
              ' Read in all dual values
57            For i = 1 To NumDualsToRead
58                Line Input #1, Line
                  ' TODO: do something here
59            Next i
              
              ' Read in all variable values
              Dim VarIndex As Long
60            For i = 1 To NumVarsToRead
61                Line Input #1, Line
62                VarIndex = VariableNLIndexToCollectionIndex(i - 1)
63                If VarIndex <= s.NumVars Then
64                    s.VarFinalValue(VarIndex) = Val(Line)
65                    s.VarCellName(VarIndex) = s.VarName(VarIndex)
66                End If
67            Next i
              
              ' Read objno suffix
68            Line Input #1, Line
              Dim solve_result_num As Long
69            solve_result_num = -1
70            If Left(Line, 5) = "objno" Then
                  Dim SplitLine() As String
71                SplitLine = Split(Line, " ")
72                If Val(SplitLine(LBound(SplitLine) + 1)) <> 0 Then
73                    RaiseGeneralError "Wrong objno"
74                End If
75                solve_result_num = Val(SplitLine(LBound(SplitLine) + 2))
76            End If
              
77            Select Case solve_result_num
              Case 0 To 99
78                s.SolveStatus = OpenSolverResult.Optimal
79                s.SolveStatusString = "Optimal"
80            Case 100 To 199
                  ' Status is `solved?`
81                Debug.Assert False
82            Case 200 To 299
83                s.SolveStatus = OpenSolverResult.Infeasible
84                s.SolveStatusString = "No Feasible Solution"
85            Case 300 To 399
86                s.SolveStatus = OpenSolverResult.Unbounded
87                s.SolveStatusString = "No Solution Found (Unbounded)"
88            Case 400 To 499
89                s.SolveStatus = OpenSolverResult.LimitedSubOptimal
90                s.SolveStatusString = "Stopped on User Limit (Time/Iterations)"
91                GetExtraInfoFromLog s
92            Case 500 To 599
93                RaiseGeneralError _
                      "There was an error while solving the model. The solver returned: " & _
                      vbNewLine & vbNewLine & SolveMessage
94            Case -1
                  ' The objno suffix wasn't there. Check SolveMessage for status
95                Debug.Assert False
96                CheckSolveMessage SolveMessage
97            Case Else
                  ' Something else
98                Debug.Assert False
99            End Select
100       End If

ExitSub:
101       Application.StatusBar = False
102       Close #1
103       If RaiseError Then RethrowError
104       Exit Sub

ErrorHandler:
105       If Not ReportError("SolverFileNL", "ReadModel_NL") Then Resume
106       RaiseError = True
107       GoTo ExitSub
End Sub

Sub CheckSolveMessage(SolveMessage As String)
          Dim LowerSolveMessage As String
1         LowerSolveMessage = LCase(SolveMessage)
          'Get the returned status code from solver.
2         If InStr(LowerSolveMessage, "optimal") > 0 Then
3             s.SolveStatus = OpenSolverResult.Optimal
4             s.SolveStatusString = "Optimal"
5         ElseIf InStr(LowerSolveMessage, "infeasible") > 0 Then
6             s.SolveStatus = OpenSolverResult.Infeasible
7             s.SolveStatusString = "No Feasible Solution"
8         ElseIf InStr(LowerSolveMessage, "unbounded") > 0 Then
9             s.SolveStatus = OpenSolverResult.Unbounded
10            s.SolveStatusString = "No Solution Found (Unbounded)"
11        ElseIf InStr(LowerSolveMessage, "interrupted on limit") > 0 Then
12            s.SolveStatus = OpenSolverResult.LimitedSubOptimal
13            s.SolveStatusString = "Stopped on User Limit (Time/Iterations)"
              ' See if we can find out which limit was hit from the log file
14            GetExtraInfoFromLog s
15        ElseIf InStr(LowerSolveMessage, "interrupted by user") > 0 Then
16            s.SolveStatus = OpenSolverResult.AbortedThruUserAction
17            s.SolveStatusString = "Stopped on Ctrl-C"
18        Else
19            If Not GetExtraInfoFromLog(s) Then
20                RaiseGeneralError _
                      "The response from the " & DisplayName(s.Solver) & " solver is not recognised. The response was: " & vbNewLine & vbNewLine & _
                      SolveMessage & vbNewLine & vbNewLine & _
                      "The " & DisplayName(s.Solver) & " command line can be found at:" & vbNewLine & _
                      SolveScriptPathName
21            End If
22        End If
End Sub

Sub CheckLog_NL(s As COpenSolver)
      ' We examine the log file if it exists to try to find errors
          
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If Not FileOrDirExists(s.LogFilePathName) Then
4             RaiseGeneralError "The solver did not create a log file. No new solution is available.", _
                                LINK_NO_SOLUTION_FILE
5         End If
          
          Dim message As String
6         Open s.LogFilePathName For Input As #3
7             message = LCase(Input$(LOF(3), 3))
8         Close #3
          
          ' We need to check > 0 explicitly, as the expression doesn't work without it
9         If Not InStr(message, LCase(s.Solver.Name)) > 0 Then
             ' Not dealing with the correct solver log, abort silently
10            GoTo ExitSub
11        End If

          ' Scan for parameter information
          Dim Key As Variant
12        For Each Key In s.SolverParameters.Keys
13            If InStr(message, LCase("Unknown keyword " & Quote(CStr(Key)))) > 0 Or _
                 InStr(message, LCase(Key & """. It is not a valid option.")) > 0 Then
14                RaiseUserError _
                      "The parameter '" & Key & "' was not recognised by the " & DisplayName(s.Solver) & " solver. " & _
                      "Please check that this name is correct, or consult the solver documentation for more information.", _
                      LINK_PARAMETER_DOCS
15            End If
16            If InStr(message, LCase("not a valid setting for Option: " & Key)) > 0 Then
17                RaiseUserError _
                      "The value specified for the parameter '" & Key & "' was invalid. " & _
                      "Please check the OpenSolver log file for a description, or consult the solver documentation for more information.", _
                      LINK_PARAMETER_DOCS
18            End If
19        Next Key

          Dim BadFunction As Variant
20        For Each BadFunction In Array("max", "min")
21            If InStr(message, LCase(BadFunction & " not implemented")) > 0 Then
22                RaiseUserError _
                      "The '" & BadFunction & "' function is not supported by the " & DisplayName(s.Solver) & " solver"
23            End If
24        Next BadFunction
          
25        If InStr(message, LCase("unknown operator")) > 0 Then
26            RaiseUserError _
                  "A function in the model is not supported by the " & DisplayName(s.Solver) & " solver. " & _
                  "This is likely to be either MIN or MAX"
27        End If

ExitSub:
28        Close #3
29        If RaiseError Then RethrowError
30        Exit Sub

ErrorHandler:
31        If Not ReportError("SolverFileNL", "CheckLog_NL") Then Resume
32        RaiseError = True
33        GoTo ExitSub
End Sub

Function GetExtraInfoFromLog(s As COpenSolver) As Boolean
          ' Checks the logs for information we can use to set the solve status
          ' This is information that isn't present in the solution file
          ' Not used to detect errors!
          
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim message As String
3         Open s.LogFilePathName For Input As #3
4         message = LCase(Input$(LOF(3), 3))
5         Close #3

          ' 1 - scan for time limit
6         If InStr(message, LCase("exiting on maximum time")) > 0 Then
7             s.SolveStatus = OpenSolverResult.LimitedSubOptimal
8             s.SolveStatusString = "Stopped on Time Limit"
9             GetExtraInfoFromLog = True
10            GoTo ExitFunction
11        End If
          ' 2 - scan for iteration limit
12        If InStr(message, LCase("exiting on maximum number of iterations")) > 0 Then
13            s.SolveStatus = OpenSolverResult.LimitedSubOptimal
14            s.SolveStatusString = "Stopped on Iteration Limit"
15            GetExtraInfoFromLog = True
16            GoTo ExitFunction
17        End If
          ' 3 - scan for infeasible. Don't look just for "infeasible", it is shown a lot even in optimal solutions
18        If InStr(message, LCase("The LP relaxation is infeasible or too expensive")) > 0 Then
19            s.SolveStatus = OpenSolverResult.Infeasible
20            s.SolveStatusString = "No Feasible Solution"
21            GetExtraInfoFromLog = True
22            GoTo ExitFunction
23        End If

ExitFunction:
24        If RaiseError Then RethrowError
25        Exit Function

ErrorHandler:
26        If Not ReportError("SolverFileNL", "CheckLogForInfo") Then Resume
27        RaiseError = True
28        GoTo ExitFunction
End Function

Function CreateSolveCommand_NL(s As COpenSolver, ScriptFilePathName As String) As String
          ' Create a script to cd to temp and run "/path/to/solver /path/to/<ModelFilePathName>"
          
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim LocalExecSolver As ISolverLocalExec, NLFileSolver As ISolverFileNL
3         Set LocalExecSolver = s.Solver
4         Set NLFileSolver = s.Solver
          
          Dim MakeOptionFile As Boolean
5         MakeOptionFile = Len(NLFileSolver.OptionFile) <> 0
             
6         CreateSolveCommand_NL = MakePathSafe(LocalExecSolver.GetExecPath()) & " " & _
                                  MakePathSafe(GetModelFilePath(s.Solver)) & " -AMPL" & _
                                  IIf(MakeOptionFile, vbNullString, ParametersToKwargs(s.SolverParameters))
          
7         If MakeOptionFile Then
              ' Create the options file in the temp folder
              Dim OptionFilePath As String
8             GetTempFilePath NLFileSolver.OptionFile, OptionFilePath
9             ParametersToOptionsFile OptionFilePath, s.SolverParameters
10        End If
          
11        CreateScriptFile ScriptFilePathName, CreateSolveCommand_NL

ExitFunction:
12        If RaiseError Then RethrowError
13        Exit Function

ErrorHandler:
14        If Not ReportError("SolverFileNL", "CreateSolveCommand_NL") Then Resume
15        RaiseError = True
16        GoTo ExitFunction
End Function

Sub CleanFiles_NL(NLFileSolver As ISolverFileNL)
1         If Len(NLFileSolver.OptionFile) <> 0 Then
              Dim OptionFilePath As String
2             GetTempFilePath NLFileSolver.OptionFile, OptionFilePath
3             DeleteFileAndVerify OptionFilePath
4         End If
End Sub
