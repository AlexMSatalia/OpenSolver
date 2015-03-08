Attribute VB_Name = "ParsedSolverNL"
Option Explicit

' This module is for writing .nl files that describe the model and solving these
' For more info on .nl file format see:
' http://citeseerx.ist.psu.edu/viewdoc/summary?doi=10.1.1.60.9659
' http://www.ampl.com/REFS/hooking2.pdf

Dim m As CModelParsed

Dim problem_name As String

Dim WriteComments As Boolean    ' Whether .nl file should include comments
Const CommentIndent = 4         ' Tracks the level of indenting in comments on nl output
Const CommentSpacing = 24       ' The column number at which nl comments begin

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

Dim VariableMap As Collection           ' A map from variable name (e.g. Test1_D4) to .nl variable index (0 to n_var - 1)
Dim VariableMapRev As Collection        ' A map from .nl variable index (0 to n_var - 1) to variable name (e.g. Test1_D4)
Public VariableIndex As Collection      ' A map from variable name (e.g. Test1_D4) to parsed variable index (1 to n_var)

Dim ConstraintMap As Collection     ' A map from constraint name (e.g. c1_Test1_D4) to .nl constraint index (0 to n_con - 1)
Dim ConstraintMapRev As Collection  ' A map from .nl constraint index (0 to n_con - 1) to constraint name (e.g. c1_Test1_D4)

Dim NonLinearConstraintTrees As Collection  ' Collection containing all non-linear constraint ExpressionTrees stored in parsed constraint order
Dim NonLinearObjectiveTrees As Collection   ' Collection containing all non-linear objective ExpressionTrees stored in parsed objective order

Dim NonLinearVars() As Boolean          ' Array of size n_var indicating whether each variable appears non-linearly in the model
Dim NonLinearConstraints() As Boolean   ' Array of size n_con indicating whether each constraint has non-linear elements

Dim ConstraintIndexToTreeIndex() As Long         ' Array of size n_con storing the parsed constraint index for each .nl constraint index
Dim VariableNLIndexToCollectionIndex() As Long   ' Array of size n_var storing the parsed variable index for each .nl variable index
Dim VariableCollectionIndexToNLIndex() As Long   ' Array of size n_var storing the .nl variable index for each parsed variable index

Dim LinearConstraints As Collection     ' Collection containing the Dictionaries for each constraint stored in parsed constraint order
Dim LinearConstants As Collection       ' Collection containing the constant (a Double) for each constraint stored in parsed constraint order
Dim LinearObjectives As Collection      ' Collection containing the Dictionaries for each objective stored in parsed objective order

Dim InitialVariableValues As Collection ' Collection containing the intital values for each variable in parsed variable index

Dim ConstraintRelations As Collection   ' Collection containing the RelationConst for each constraint stored in parsed constraint order

Dim ObjectiveCells As Collection    ' Collection containing the objective cells for each objective stored in parsed objective order
Dim ObjectiveSenses As Collection   ' Collection containing the objective sense for each objective stored in parsed objective order

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
7458      GetVariableNLIndex = VariableCollectionIndexToNLIndex(Index)
End Property

' Creates .nl file and solves model
Function SolveModelParsed_NL(ModelFilePathName As String, model As CModelParsed, s As COpenSolverParsed, SolveOptions As SolveOptionsType, SolveRelaxation As Boolean, Optional ShouldWriteComments As Boolean = True)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          Application.EnableCancelKey = xlErrorHandler

7459      Set m = model
          
7460      WriteComments = ShouldWriteComments
          
          ' =============================================================
          ' Process model for .nl output
          ' All Module-level variables required for .nl output should be set in this step
          ' No modification to these variables should be done while writing the .nl file
          ' =============================================================
          
7462      InitialiseModelStats
          
7464      If n_obj = 0 And n_con = 0 Then
7465          Err.Raise OpenSolver_BuildError, "The model has no constraints that depend on the adjustable cells, and has no objective. There is nothing for the solver to do."
7466      End If
          
7467      CreateVariableIndex
          
7468      ProcessFormulae
7469      ProcessObjective
          
7470      MakeVariableMap SolveRelaxation
7471      MakeConstraintMap
          
          ' =============================================================
          ' Write output files
          ' =============================================================
          
          ' Create supplementary outputs
7472      OutputColFile
7473      OutputRowFile
          
          ' Write .nl file
7474      Open ModelFilePathName For Output As #1
          
7475      WriteToFile 1, MakeHeader(), AbortIfBlank:=True
7477      WriteToFile 1, MakeCBlocks(), AbortIfBlank:=True
7480      WriteToFile 1, MakeOBlocks(), AbortIfBlank:=True
7483      'WriteToFile 1, MakeDBlock(), AbortIfBlank:=True
7485      WriteToFile 1, MakeXBlock(), AbortIfBlank:=True
7487      WriteToFile 1, MakeRBlock(), AbortIfBlank:=True
7489      WriteToFile 1, MakeBBlock(), AbortIfBlank:=True
7491      WriteToFile 1, MakeKBlock(), AbortIfBlank:=True
7494      WriteToFile 1, MakeJBlocks(), AbortIfBlank:=True
7497      WriteToFile 1, MakeGBlocks(), AbortIfBlank:=True

7499      Close #1
          
          ' =============================================================
          ' Solve model using chosen solver
          ' =============================================================
          
7501      UpdateStatusBar "OpenSolver: Solving .nl model file", True
          
          Dim SolutionFilePathName As String
7502      SolutionFilePathName = SolutionFilePath(m.Solver)
          
7503      DeleteFileAndVerify SolutionFilePathName

          Dim ExternalSolverPathName As String
7504      ExternalSolverPathName = CreateSolveScriptParsed(m.Solver, ModelFilePathName, SolveOptions)
                   
          Dim logCommand As String, logFileName As String
7505      logFileName = "log1.tmp"
7506      logCommand = MakePathSafe(GetTempFilePath(logFileName))
                        
          Dim ExecutionCompleted As Boolean
7507      ExternalSolverPathName = MakePathSafe(ExternalSolverPathName)
                    
          Dim exeResult As Long
7508      ExecutionCompleted = RunExternalCommand(ExternalSolverPathName, logCommand, IIf(s.GetShowIterationResults, SW_SHOWNORMAL, SW_HIDE), True, exeResult) ' Run solver, waiting for completion
7513      If exeResult <> 0 Then
7515          Err.Raise Number:=OpenSolver_SolveError, Description:="The " & m.Solver & " solver did not complete, but aborted with the error code " & exeResult & "." & vbCrLf & vbCrLf & "The last log file can be viewed under the OpenSolver menu and may give you more information on what caused this error."
7516      End If
          
          ' =============================================================
          ' Read results from solution file
          ' =============================================================
7518      UpdateStatusBar "OpenSolver: Reading .nl solution", True

          Dim solutionLoaded As Boolean, errorString As String
7519      solutionLoaded = ReadModel_NL(SolutionFilePathName, errorString, s)
7521      If errorString <> "" Then
7522          Err.Raise Number:=OpenSolver_SolveError, Description:=errorString
7523      ElseIf Not solutionLoaded Then 'read error
7524          SolveModelParsed_NL = False
7525          GoTo ExitFunction
7526      End If

7527      SolveModelParsed_NL = True
          
ExitFunction:
7529      Application.StatusBar = False
7530      Close #1
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "SolveModelParsed_NL") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Private Sub InitialiseModelStats()
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          ' Number of actual variables is the number of adjustable cells in the Solver model
7547      numActualVars = m.AdjustableCells.Count
          ' Number of fake variables is the number of formulae equations we have created
7548      numFakeVars = m.Formulae.Count
          ' Number of actual constraints is the number of constraints in the Solver model
7549      numActualCons = m.LHSKeys.Count
          ' Number of fake constraints is the number of formulae equations we have created
7550      numFakeCons = m.Formulae.Count
          
          ' Divide the actual constraints into equalities and inequalities (ranges)
          Dim i As Long
7551      numActualEqs = 0
7552      numActualRanges = 0
7553      For i = 1 To numActualCons
7554          UpdateStatusBar "OpenSolver: Creating .nl file. Counting constraints: " & i & "/" & numActualCons & ". "
              
7558          If m.Rels(i) = RelationConsts.RelationEQ Then
7559              numActualEqs = numActualEqs + 1
7560          Else
7561              numActualRanges = numActualRanges + 1
7562          End If
7563      Next i
          
          ' ===============================================================================
          ' Initialise the ASL variables - see definitions for explanation of each variable
          
          ' Model statistics for line #1
7564      problem_name = "'Sheet=" + m.SolverModelSheet.Name + "'"
          
          ' Model statistics for line #2
7565      n_var = numActualVars + numFakeVars
7566      n_con = numActualCons + numFakeCons
7567      n_obj = 0
7568      nranges = numActualRanges
7569      n_eqn_ = numActualEqs + numFakeCons     ' All fake formulae constraints are equalities
7570      n_lcon = 0
          
          ' Model statistics for line #3
7571      nlc = 0
7572      nlo = 0
          
          ' Model statistics for line #4
7573      nlnc = 0
7574      lnc = 0
          
          ' Model statistics for line #5
7575      nlvc = 0
7576      nlvo = 0
7577      nlvb = 0
          
          ' Model statistics for line #6
7578      nwv_ = 0
7579      nfunc_ = 0
7580      arith = 0
7581      flags = 0
          
          ' Model statistics for line #7
7582      nbv = 0
7583      niv = 0
7584      nlvbi = 0
7585      nlvci = 0
7586      nlvoi = 0
          
          ' Model statistics for line #8
7587      nzc = 0
7588      nzo = 0
          
          ' Model statistics for line #9
7589      maxrownamelen_ = 0
7590      maxcolnamelen_ = 0
          
          ' Model statistics for line #10
7591      comb = 0
7592      comc = 0
7593      como = 0
7594      comc1 = 0
7595      como1 = 0

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "InitialiseModelStats") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

' Creates map from variable name (e.g. Test1_D4) to parsed variable index (1 to n_var)
Private Sub CreateVariableIndex()
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

7596      Set VariableIndex = New Collection
7597      Set InitialVariableValues = New Collection
          Dim c As Range, cellName As String, i As Long
          
          ' First read in actual vars
7598      i = 1
7599      For Each c In m.AdjustableCells
7600          UpdateStatusBar "OpenSolver: Creating .nl file. Counting variables: " & i & "/" & numActualVars & ". "
              
7604          cellName = ConvertCellToStandardName(c)
              
              ' Update variable maps
7605          VariableIndex.Add i, cellName
              
              ' Update initial values
7606          InitialVariableValues.Add CDbl(c)
              
7607          i = i + 1
7608      Next c
          
          ' Next read in fake formulae vars
7609      For i = 1 To numFakeVars
7610          UpdateStatusBar "OpenSolver: Creating .nl file. Counting formulae variables: " & i & "/" & numFakeVars & ". "
              
7614          cellName = m.Formulae(i).strAddress
              
              ' Update variable maps
7615          VariableIndex.Add i + numActualVars, cellName
7616      Next i

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "CreateVariableIndex") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

' Creates maps from variable name (e.g. Test1_D4) to .nl variable index (0 to n_var - 1) and vice-versa, and
' maps from parsed variable index to .nl variable index and vice-versa
Private Sub MakeVariableMap(SolveRelaxation As Boolean)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          ' Create index of variable names in parsed variable order
          Dim CellNames As New Collection
          
          ' Actual variables
          Dim c As Range, cellName As String, i As Long
7617      i = 1
7618      For Each c In m.AdjustableCells
7619          UpdateStatusBar "OpenSolver: Creating .nl file. Classifying variables: " & i & "/" & numActualVars & ". "
              
7623          i = i + 1
7624          cellName = ConvertCellToStandardName(c)
7625          CellNames.Add cellName
7626      Next c
          
          ' Formulae variables
7627      For i = 1 To m.Formulae.Count
7628          UpdateStatusBar "OpenSolver: Creating .nl file. Classifying formulae variables: " & i & "/" & numFakeVars & ". "
              
7632          cellName = m.Formulae(i).strAddress
7633          CellNames.Add cellName
7634      Next i
          
          '==============================================
          ' Get binary and integer variables from model
          
          ' Get integer variables
          Dim IntegerVars() As Boolean
7635      ReDim IntegerVars(n_var)
7636      If Not m.IntegerCells Is Nothing Then
7638          UpdateStatusBar "OpenSolver: Creating .nl file. Finding integer variables"
7637          For Each c In m.IntegerCells
7640              cellName = ConvertCellToStandardName(c)
7641              IntegerVars(VariableIndex(cellName)) = True
7642          Next c
7643      End If
          
          ' Get binary variables
7644      ReDim BinaryVars(n_var)
7645      If Not m.BinaryCells Is Nothing Then
7647          UpdateStatusBar "OpenSolver: Creating .nl file. Finding binary variables"
7646          For Each c In m.BinaryCells
7649              cellName = ConvertCellToStandardName(c)
7650              BinaryVars(VariableIndex(cellName)) = True
7652          Next c
7653      End If
          
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
          
7654      For i = 1 To n_var
7655          UpdateStatusBar "OpenSolver: Creating .nl file. Sorting variables: " & i & "/" & n_var & ". "
              
7659          If NonLinearVars(i) Then
7660              If IntegerVars(i) Or BinaryVars(i) Then
7661                  NonLinearInteger.Add i
7662              Else
7663                  NonLinearContinuous.Add i
7664              End If
7665          Else
7666              If BinaryVars(i) Then
7667                  LinearBinary.Add i
7668              ElseIf IntegerVars(i) Then
7669                  LinearInteger.Add i
7670              Else
7671                  LinearContinuous.Add i
7672              End If
7673          End If
7674      Next i
          
          ' ==============================================
          ' Add variables to the variable map in the required order
7675      ReDim VariableNLIndexToCollectionIndex(n_var)
7676      ReDim VariableCollectionIndexToNLIndex(n_var)
7677      Set VariableMap = New Collection
7678      Set VariableMapRev = New Collection
          
          Dim Index As Long, var As Long
7679      Index = 0
          
          ' We loop through the variables and arrange them in the required order:
          '     1st - non-linear continuous
          '     2nd - non-linear integer
          '     3rd - linear arcs (N/A)
          '     4th - other linear
          '     5th - binary
          '     6th - other integer
          
          ' Non-linear continuous
7680      For i = 1 To NonLinearContinuous.Count
7681          UpdateStatusBar "OpenSolver: Creating .nl file. Outputting non-linear continuous vars"
              
7685          var = NonLinearContinuous(i)
7686          AddVariable CellNames(var), Index, var
7687      Next i
          
          ' Non-linear integer
7688      For i = 1 To NonLinearInteger.Count
7689          UpdateStatusBar "OpenSolver: Creating .nl file. Outputting non-linear integer vars"
              
7693          var = NonLinearInteger(i)
7694          AddVariable CellNames(var), Index, var
7695      Next i
          
          ' Linear continuous
7696      For i = 1 To LinearContinuous.Count
7697          UpdateStatusBar "OpenSolver: Creating .nl file. Outputting linear continuous vars"

7701          var = LinearContinuous(i)
7702          AddVariable CellNames(var), Index, var
7703      Next i
          
          ' Linear binary
7704      For i = 1 To LinearBinary.Count
7705          UpdateStatusBar "OpenSolver: Creating .nl file. Outputting linear binary vars"
              
7709          var = LinearBinary(i)
7710          AddVariable CellNames(var), Index, var
7711      Next i
          
          ' Linear integer
7712      For i = 1 To LinearInteger.Count
7713          UpdateStatusBar "OpenSolver: Creating .nl file. Outputting linear integer vars"

7717          var = LinearInteger(i)
7718          AddVariable CellNames(var), Index, var
7719      Next i
          
          ' ==============================================
          ' Update model stats
          If Not SolveRelaxation Then
7720          nbv = LinearBinary.Count
7721          niv = LinearInteger.Count
7722          nlvci = NonLinearInteger.Count
          End If

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "MakeVariableMap") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

' Adds a variable to all the variable maps with:
'   variable name:            cellName
'   .nl variable index:       index
'   parsed variable index:    i
Private Sub AddVariable(cellName As String, Index As Long, i As Long)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          ' Update variable maps
7723      VariableMap.Add CStr(Index), cellName
7724      VariableMapRev.Add cellName, CStr(Index)
7725      VariableNLIndexToCollectionIndex(Index) = i
7726      VariableCollectionIndexToNLIndex(i) = Index
          
          ' Update max length of variable name
7727      If Len(cellName) > maxcolnamelen_ Then
7728         maxcolnamelen_ = Len(cellName)
7729      End If
          
          ' Increase index for the next variable
7730      Index = Index + 1

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "AddVariable") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

' Creates maps from constraint name (e.g. c1_Test1_D4) to .nl constraint index (0 to n_con - 1) and vice-versa, and
' map from .nl constraint index to parsed constraint index
Private Sub MakeConstraintMap()
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

7731      Set ConstraintMap = New Collection
7732      Set ConstraintMapRev = New Collection
7733      ReDim ConstraintIndexToTreeIndex(n_con)
          
          Dim Index As Long, i As Long, cellName As String
7734      Index = 0
          
          ' We loop through the constraints and arrange them in the required order:
          '     1st - non-linear
          '     2nd - non-linear network (N/A)
          '     3rd - linear network (N/A)
          '     4th - linear
          
          ' Non-linear constraints
7735      For i = 1 To n_con
7736          UpdateStatusBar "OpenSolver: Creating .nl file. Outputting non-linear constraints " & i & "/" & n_con

7740          If NonLinearConstraints(i) Then
                  ' Actual constraints
7741              If i <= numActualCons Then
7742                  cellName = "c" & i & "_" & m.LHSKeys(i)
                  ' Formulae constraints
7743              ElseIf i <= numActualCons + numFakeCons Then
7744                  cellName = "f" & i & "_" & m.Formulae(i - numActualCons).strAddress
7745              Else
7746                  cellName = "seek_obj_" & ConvertCellToStandardName(m.ObjectiveCell)
7747              End If
7748              AddConstraint cellName, Index, i
7749          End If
7750      Next i
          
          ' Linear constraints
7751      For i = 1 To n_con
7752          UpdateStatusBar "OpenSolver: Creating .nl file. Outputting linear constraints " & i & "/" & n_con

7756          If Not NonLinearConstraints(i) Then
                  ' Actual constraints
7757              If i <= numActualCons Then
7758                  cellName = "c" & i & "_" & m.LHSKeys(i)
                  ' Formulae constraints
7759              ElseIf i <= numActualCons + numFakeCons Then
7760                  cellName = "f" & i & "_" & m.Formulae(i - numActualCons).strAddress
7761              Else
7762                  cellName = "seek_obj_" & ConvertCellToStandardName(m.ObjectiveCell)
7763              End If
7764              AddConstraint cellName, Index, i
7765          End If
7766      Next i

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "MakeConstraintMap") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

' Adds a constraint to all the constraint maps with:
'   constraint name:          cellName
'   .nl constraint index:     index
'   parsed constraint index:  i
Private Sub AddConstraint(cellName As String, Index As Long, i As Long)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          ' Update constraint maps
7767      ConstraintMap.Add Index, cellName
7768      ConstraintMapRev.Add cellName, CStr(Index)
7769      ConstraintIndexToTreeIndex(Index) = i
          
          ' Update max length
7770      If Len(cellName) > maxrownamelen_ Then
7771         maxrownamelen_ = Len(cellName)
7772      End If
          
          ' Increase index for next constraint
7773      Index = Index + 1

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "AddConstraint") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

' Processes all constraint formulae into the .nl model formats
Private Sub ProcessFormulae()
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

7774      Set NonLinearConstraintTrees = New Collection
7775      Set LinearConstraints = New Collection
7776      Set LinearConstants = New Collection
7777      Set ConstraintRelations = New Collection
          
7778      ReDim NonLinearVars(n_var)
7779      ReDim NonLinearConstraints(n_con)
7780      ReDim NonZeroConstraintCount(n_var)
          
          ' Loop through all constraints and process each
          Dim i As Long
7781      For i = 1 To numActualCons
7782          UpdateStatusBar "OpenSolver: Processing formulae into expression trees... " & i & "/" & n_con & " formulae."
7786          ProcessSingleFormula m.RHSKeys(i), m.LHSKeys(i), m.Rels(i)
7787      Next i
          
7788      For i = 1 To numFakeCons
7789          UpdateStatusBar "OpenSolver: Processing formulae into expression trees... " & i + numActualCons & "/" & n_con & " formulae."
7793          ProcessSingleFormula m.Formulae(i).strFormulaParsed, m.Formulae(i).strAddress, RelationConsts.RelationEQ
7794      Next i
          
ExitSub:
7795      Application.StatusBar = False
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "ProcessFormulae") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

' Processes a single constraint into .nl format. We require:
'     - a non-linear ExpressionTree for all non-linear parts of the equation
'     - a linear Dictionary for the linear parts of the equation
'     - a constant Double for the constant part of the equation
' We also use the results of processing to update some of the model statistics
Private Sub ProcessSingleFormula(RHSExpression As String, LHSVariable As String, Relation As RelationConsts)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          ' Convert the string formula into an ExpressionTree object
          Dim Tree As ExpressionTree
7796      Set Tree = ConvertFormulaToExpressionTree(RHSExpression)
          
          ' The .nl file needs a linear coefficient for every variable in the constraint - non-linear or otherwise
          ' We need a list of all variables in this constraint so that we can know to include them in the linear part of the constraint.
          Dim constraint As New Dictionary
7798      Tree.ExtractVariables constraint

          Dim LinearTrees As New Collection
7799      Tree.MarkLinearity
          
          ' Constants in .nl expression trees must be as simple as possible.
          ' We cannot have constant * constant in the tree, it must be replaced with a single constant
          ' We need to evalulate and pull up all constants that we can.
7800      Tree.PullUpConstants
          
          ' Remove linear terms from non-linear trees
7801      Tree.PruneLinearTrees LinearTrees
          
          ' Process linear trees to separate constants and variables
          Dim constant As Double
7802      constant = 0
          Dim j As Long
7803      For j = 1 To LinearTrees.Count
7804          LinearTrees(j).ConvertLinearTreeToConstraint constraint, constant
7805      Next j
          
          ' Check that our variable LHS exists
7806      If Not TestKeyExists(VariableIndex, LHSVariable) Then
              ' We must have a constant formula as the LHS, we can evaluate and merge with the constant
              ' We are bringing the constant to the RHS so must subtract
7807          If Not left(LHSVariable, 1) = "=" Then LHSVariable = "=" & LHSVariable
7808          constant = constant - Application.Evaluate(LHSVariable)
7809      Else
              ' Our constraint has a single term on the LHS and a formulae on the right.
              ' The single LHS term needs to be included in the expression as a linear term with coefficient 1
              ' Move variable from the LHS to the linear constraint with coefficient -1
              Dim VarKey As Long
              VarKey = VariableIndex(LHSVariable)
              If constraint.Exists(VarKey) Then
                  constraint.Item(VarKey) = constraint.Item(VarKey) - 1
              Else
7810              constraint.Add Key:=VariableIndex(LHSVariable), Item:=-1
              End If
7812      End If
          
          ' Keep constant on RHS and take everything else to LHS.
          ' Need to flip all sign of elements in this constraint

          ' Flip all coefficients in linear constraint
7821      InvertCoefficients constraint

          ' Negate non-linear tree
7822      Set Tree = Tree.Negate
          
          ' Save results of processing
7824      NonLinearConstraintTrees.Add Tree
7825      LinearConstraints.Add constraint
7826      LinearConstants.Add constant
7827      ConstraintRelations.Add Relation
          
          ' Mark any non-linear variables that we haven't seen before by extracting any remaining variables from the non-linear tree.
          ' Any variable still present in the constraint must be part of the non-linear section
          Dim TempConstraint As New Dictionary, var As Variant
7829      Tree.ExtractVariables TempConstraint
7830      For Each var In TempConstraint.Keys
7831          If Not NonLinearVars(var) Then
7832              NonLinearVars(var) = True
7833              nlvc = nlvc + 1
7834          End If
7835      Next var
          
          ' Remove any zero coefficients from the linear constraint if the variable is not in the non-linear tree
7836      For Each var In constraint.Keys
7837          If Not NonLinearVars(var) And constraint.Item(var) = 0 Then
7838              constraint.Remove var
7839          End If
7840      Next var
          
          ' Increase count of non-linear constraints if the non-linear tree is non-empty
          ' An empty tree has a single "0" node
7841      If Tree.NodeText <> "0" Then
              Dim ConstraintIndex As Long
7842          ConstraintIndex = NonLinearConstraintTrees.Count
7843          NonLinearConstraints(ConstraintIndex) = True
7844          nlc = nlc + 1
7845      End If
          
          ' Update jacobian counts using the linear variables present
7846      For Each var In constraint.Keys
              ' The nl documentation says that the jacobian counts relate to "the numbers of nonzeros in the first n var - 1 columns of the Jacobian matrix"
              ' This means we should increases the count for any variable that is present and has a non-zero coefficient
              ' However, .nl files generated by AMPL seem to increase the count for any present variable, even if the coefficient is zero (i.e. non-linear variables)
              ' We adopt the AMPL behaviour as it gives faster solve times (see Test 40). Swap the If conditions below to change this behaviour
              
              'If constraint.Coefficient(j) <> 0 Then
7847          NonZeroConstraintCount(CLng(var)) = NonZeroConstraintCount(CLng(var)) + 1
7849          nzc = nzc + 1
              'End If
7851      Next var

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "ProcessSingleFormula") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Sub InvertCoefficients(ByRef constraint As Dictionary)
    Dim Key As Variant
    For Each Key In constraint.Keys
        constraint.Item(Key) = -constraint.Item(Key)
    Next Key
End Sub

' Process objective function into .nl format. We require the same as for constraints:
'     - a non-linear ExpressionTree for all non-linear parts of the equation
'     - a linear Dictionary for the linear parts of the equation
'     - a constant Double for the constant part of the equation
Private Sub ProcessObjective()
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

7852      Set NonLinearObjectiveTrees = New Collection
7853      Set ObjectiveSenses = New Collection
7854      Set ObjectiveCells = New Collection
7855      Set LinearObjectives = New Collection
          
          ' =======================================================
          ' Currently just one objective - a single linear variable
          ' We could move to multiple objectives if OpenSolver supported this
          
7856      If m.ObjectiveSense = TargetObjective Then
              ' Instead of adding an objective, we add a constraint
7857          ProcessSingleFormula m.ObjectiveTargetValue, ConvertCellToStandardName(m.ObjectiveCell), RelationEQ
7858          n_con = n_con + 1
7859          ReDim Preserve NonLinearConstraints(n_con)
7860      ElseIf m.ObjectiveCell Is Nothing Then
              ' Do nothing is objective is missing
7861      Else
7862          ObjectiveCells.Add ConvertCellToStandardName(m.ObjectiveCell)
7863          ObjectiveSenses.Add m.ObjectiveSense
              
              ' Objective non-linear constraint tree is empty
7864          NonLinearObjectiveTrees.Add CreateTree("0", ExpressionTreeNodeType.ExpressionTreeNumber)
              
              ' Objective has a single linear term - the objective variable with coefficient 1
              Dim Objective As New Dictionary
7866          Objective.Add Key:=VariableIndex(ObjectiveCells(1)), Item:=1
              
              ' Save results
7868          LinearObjectives.Add Objective
              
              ' Track non-zero jacobian count in objective
7869          nzo = nzo + 1
              
              ' Adjust objective count
7870          n_obj = n_obj + 1
7871      End If

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "ProcessObjective") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

' Writes header block for .nl file. This contains the model statistics
Private Function MakeHeader() As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim Header As String
7872      Header = ""
          
7873      AddNewLine Header, "g3 1 1 0", "problem " & problem_name
7874      AddNewLine Header, " " & n_var & " " & n_con & " " & n_obj & " " & nranges & " " & n_eqn_ & " " & n_lcon, "vars, constraints, objectives, ranges, eqns"
7875      AddNewLine Header, " " & nlc & " " & nlo, "nonlinear constraints, objectives"
7876      AddNewLine Header, " " & nlnc & " " & lnc, "network constraints: nonlinear, linear"
7877      AddNewLine Header, " " & nlvc & " " & nlvo & " " & nlvb, "nonlinear vars in constraints, objectives, both"
7878      AddNewLine Header, " " & nwv_ & " " & nfunc_ & " " & arith & " " & flags, "linear network variables; functions; arith, flags"
7879      AddNewLine Header, " " & nbv & " " & niv & " " & nlvbi & " " & nlvci & " " & nlvoi, "discrete variables: binary, integer, nonlinear (b,c,o)"
7880      AddNewLine Header, " " & nzc & " " & nzo, "nonzeros in Jacobian, gradients"
7881      AddNewLine Header, " " & maxrownamelen_ & " " & maxcolnamelen_, "max name lengths: constraints, variables"
7882      AddNewLine Header, " " & comb & " " & comc & " " & como & " " & comc1 & " " & como1, "common exprs: b,c,o,c1,o1"
          
7883      MakeHeader = StripTrailingNewline(Header)

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "MakeHeader") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

' Writes C blocks for .nl file. These describe the non-linear parts of each constraint.
Private Function MakeCBlocks() As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim Block As String
7884      Block = ""
          
          Dim i As Long
7885      For i = 1 To n_con
7886          UpdateStatusBar "OpenSolver: Creating .nl file. Writing non-linear constraints " & i & "/" & n_con

              ' Add block header for the constraint
7890          AddNewLine Block, "C" & i - 1, "CONSTRAINT NON-LINEAR SECTION " + ConstraintMapRev(CStr(i - 1))
              
              ' Add expression tree
7892          Block = Block + NonLinearConstraintTrees(ConstraintIndexToTreeIndex(i - 1)).ConvertToNL(CommentIndent)
7893      Next i
          
7894      MakeCBlocks = StripTrailingNewline(Block)

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "MakeCBlocks") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

' Writes O blocks for .nl file. These describe the non-linear parts of each objective.
Private Function MakeOBlocks() As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim Block As String
7895      Block = ""
          
          Dim i As Long
7896      For i = 1 To n_obj
              ' Add block header for the objective
7897          AddNewLine Block, "O" & i - 1 & " " & ConvertObjectiveSenseToNL(ObjectiveSenses(i)), "OBJECTIVE NON-LINEAR SECTION " & ObjectiveCells(i)
              
              ' Add expression tree
7899          Block = Block + NonLinearObjectiveTrees(i).ConvertToNL(CommentIndent)
7900      Next i
          
7901      MakeOBlocks = StripTrailingNewline(Block)

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "MakeOBlocks") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

' Writes D block for .nl file. This contains the initial guess for dual variables.
' We don't use this, so just set them all to zero
Private Function MakeDBlock() As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim Block As String
7902      Block = ""
          
          ' Add block header
7903      AddNewLine Block, "d" & n_con, "INITIAL DUAL GUESS"
          
          ' Set duals to zero for all constraints
          Dim i As Long
7904      For i = 1 To n_con
7905          UpdateStatusBar "OpenSolver: Creating .nl file. Writing initial duals " & i & "/" & n_con
              
7909          AddNewLine Block, i - 1 & " 0", "    " & ConstraintMapRev(CStr(i - 1)) & " = " & 0
7910      Next i
          
7911      MakeDBlock = StripTrailingNewline(Block)

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "MakeDBlock") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

' Writes X block for .nl file. This contains the initial guess for primal variables
Private Function MakeXBlock() As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim Block As String
7912      Block = ""
          
          ' Add block header
7913      AddNewLine Block, "x" & n_var, "INITIAL PRIMAL GUESS"

          ' Loop through the variables in .nl variable order
          Dim i As Long, initial As Double, VariableIndex As Long
7914      For i = 1 To n_var
7915          UpdateStatusBar "OpenSolver: Creating .nl file. Writing initial values " & i & "/" & n_var
              
7919          VariableIndex = VariableNLIndexToCollectionIndex(i - 1)
              
              ' Get initial values
7920          If VariableIndex <= numActualVars Then
                  ' Actual variables - use the value in the actual cell
7921              initial = InitialVariableValues(VariableIndex)
7922          Else
                  ' Formulae variables - use the initial value saved in the CFormula instance
7923              initial = CDbl(m.Formulae(VariableIndex - numActualVars).initialValue)
7924          End If
7925          AddNewLine Block, i - 1 & " " & StrEx_NL(initial), "    " & VariableMapRev(CStr(i - 1)) & " = " & initial
7926      Next i
          
7927      MakeXBlock = StripTrailingNewline(Block)

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "MakeXBlock") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

' Writes R block for .nl file. This contains the constant values for each constraint and the relation type
Private Function MakeRBlock() As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          If n_con = 0 Then GoTo ExitFunction

          Dim Block As String
7928      Block = ""
           ' Add block header
7929      AddNewLine Block, "r", "CONSTRAINT BOUNDS"
          
          ' Apply bounds according to the relation type
          Dim i As Long, BoundType As Long, Comment As String, bound As Double
7930      For i = 1 To n_con
7931          UpdateStatusBar "OpenSolver: Creating .nl file. Writing constraint bounds " & i & "/" & n_con
              
7935          bound = LinearConstants(ConstraintIndexToTreeIndex(i - 1))
7936          ConvertConstraintToNL ConstraintRelations(ConstraintIndexToTreeIndex(i - 1)), BoundType, Comment
7937          AddNewLine Block, BoundType & " " & StrEx_NL(bound), "    " & ConstraintMapRev(CStr(i - 1)) & Comment & bound
7938      Next i
          
7939      MakeRBlock = StripTrailingNewline(Block)

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "MakeRBlock") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

' Writes B block for .nl file. This contains the variable bounds
Private Function MakeBBlock() As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim Block As String
7940      Block = ""
          
          ' Write block header
7941      AddNewLine Block, "b", "VARIABLE BOUNDS"
          
          Dim i As Long, bound As String, Comment As String, VariableIndex As Long, VarName As String
7942      For i = 1 To n_var
7943          UpdateStatusBar "OpenSolver: Creating .nl file. Writing variable bounds " & i & "/" & n_var
              
7947          VariableIndex = VariableNLIndexToCollectionIndex(i - 1)
7948          Comment = "    " & VariableMapRev(CStr(i - 1))
           
7949          If VariableIndex <= numActualVars Then
                  If BinaryVars(VariableIndex) Then
                    bound = "0 0 1"
                    Comment = Comment & " IN [0, 1]"
                  ' Real variables, use actual bounds
7950              ElseIf m.AssumeNonNegative Then
7951                  VarName = m.AdjCellNames(CLng(VariableIndex))
7952                  If TestKeyExists(m.VarLowerBounds, VarName) And Not BinaryVars(VariableIndex) Then
7953                      bound = "3"
7954                      Comment = Comment & " FREE"
7955                  Else
7956                      bound = "2 0"
7957                      Comment = Comment & " >= 0"
7958                  End If
7959              Else
7960                  bound = "3"
7961                  Comment = Comment & " FREE"
7962              End If
7963          Else
                  ' Fake formulae variables - no bounds
7964              bound = "3"
7965              Comment = Comment & " FREE"
7966          End If
7967          AddNewLine Block, bound, Comment
7968      Next i
          
7969      MakeBBlock = StripTrailingNewline(Block)

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "MakeBBlock") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

' Writes K block for .nl file. This contains the cumulative count of non-zero jacobian entries for the first n-1 variables
Private Function MakeKBlock() As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          If n_var = 0 Then GoTo ExitFunction

          Dim Block As String
7970      Block = ""

          ' Add block header
7971      AddNewLine Block, "k" & n_var - 1, "NUMBER OF JACOBIAN ENTRIES (CUMULATIVE) FOR FIRST " & n_var - 1 & " VARIABLES"
          
          ' Loop through first n_var - 1 variables and add the non-zero count to the running total
          Dim i As Long, total As Long
7972      total = 0
7973      For i = 1 To n_var - 1
7974          UpdateStatusBar "OpenSolver: Creating .nl file. Writing jacobian counts " & i & "/" & n_var - 1
              
7978          total = total + NonZeroConstraintCount(VariableNLIndexToCollectionIndex(i - 1))
7979          AddNewLine Block, CStr(total), "    Up to " & VariableMapRev(CStr(i - 1)) & ": " & CStr(total) & " entries in Jacobian"
7980      Next i
          
7981      MakeKBlock = StripTrailingNewline(Block)

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "MakeKBlock") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

' Writes J blocks for .nl file. These contain the linear part of each constraint
Private Function MakeJBlocks() As String
          Dim Block As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

7982      Block = ""
          
          Dim ConstraintElements() As String
          Dim CommentElements() As String
          
          Dim i As Long, TreeIndex As Long, VariableIndex As Long
7983      For i = 1 To n_con
7984          UpdateStatusBar "OpenSolver: Creating .nl file. Writing linear constraints " & i & "/" & n_con
              
7988          TreeIndex = ConstraintIndexToTreeIndex(i - 1)
          
              ' Make header
7989          AddNewLine Block, "J" & i - 1 & " " & LinearConstraints(TreeIndex).Count, "CONSTRAINT LINEAR SECTION " & ConstraintMapRev(i)
              
7990          ReDim ConstraintElements(n_var)
7991          ReDim CommentElements(n_var)
              
              ' We need variables output in .nl order
              ' First we collect all the constraint elements for the variables that are present
              Dim j As Long
7992          For j = 1 To n_var
7993              If LinearConstraints(TreeIndex).Exists(j) Then
7994                  VariableIndex = VariableCollectionIndexToNLIndex(j)
7995                  ConstraintElements(VariableIndex) = VariableIndex & " " & StrEx_NL(LinearConstraints(TreeIndex).Item(j))
7996                  CommentElements(VariableIndex) = "    + " & LinearConstraints(TreeIndex).Item(j) & " * " & VariableMapRev(CStr(VariableIndex))
7997              End If
7998          Next j
              
              ' Output the constraint elements to the J Block in .nl variable order
7999          For j = 0 To n_var - 1
8000              If ConstraintElements(j) <> "" Then
8001                  AddNewLine Block, ConstraintElements(j), CommentElements(j)
8002              End If
8003          Next j
              
8004      Next i
          
8005      MakeJBlocks = StripTrailingNewline(Block)

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "MakeJBlocks") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

' Writes the G blocks for .nl file. These contain the linear parts of each objective
Private Function MakeGBlocks() As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim Block As String
8006      Block = ""
          
          Dim i As Long
8007      For i = 1 To n_obj
              ' Make header
8008          AddNewLine Block, "G" & i - 1 & " " & LinearObjectives(i).Count, "OBJECTIVE LINEAR SECTION " & ObjectiveCells(i)
              
              ' This loop is not in the right order (see J blocks)
              ' Since the objective only containts one variable, the output will still be correct without reordering
              Dim j As Long, VariableIndex As Long
8009          For j = 1 To n_var
8010              If LinearObjectives(i).Exists(j) Then
8011                  VariableIndex = VariableCollectionIndexToNLIndex(j)
8012                  AddNewLine Block, VariableIndex & " " & StrEx_NL(LinearObjectives(i).Item(j)), "    + " & LinearObjectives(i).Item(j) & " * " & VariableMapRev(CStr(VariableIndex))
8013              End If
8014          Next j
8015      Next i
          
8016      MakeGBlocks = StripTrailingNewline(Block)

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "MakeGBlocks") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

' Writes the .col summary file. This contains the variable names listed in .nl order
Private Sub OutputColFile()
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim ColFilePathName As String
8017      ColFilePathName = GetTempFilePath("model.col")
          
8018      DeleteFileAndVerify ColFilePathName

8020      Open ColFilePathName For Output As #2
          
8022      UpdateStatusBar "OpenSolver: Creating .nl file. Writing col file"
          Dim var As Variant
8021      For Each var In VariableMap
8024          WriteToFile 2, VariableMapRev(var)
8025      Next var
          
ExitSub:
8026      Close #2
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "OutputColFile") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

' Writes the .row summary file. This contains the constraint names listed in .nl order
Private Sub OutputRowFile()
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim RowFilePathName As String
8030      RowFilePathName = GetTempFilePath("model.row")
          
8031      DeleteFileAndVerify RowFilePathName

8033      Open RowFilePathName For Output As #3
          
8035      UpdateStatusBar "OpenSolver: Creating .nl file. Writing con file"
          Dim con As Variant
8034      For Each con In ConstraintMap
8036          DoEvents
8037          WriteToFile 3, ConstraintMapRev(CStr(con))
8038      Next con

ExitSub:
          Close #3
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "OutputRowFile") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Private Sub OutputOptionsFile(OptionsFilePath As String, SolveOptions As SolveOptionsType)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          DeleteFileAndVerify OptionsFilePath
          
          Open OptionsFilePath For Output As #4
          Print #4, "iteration_limit " & str(SolveOptions.MaxIterations)
          Print #4, "allowable_fraction_gap " & StrEx_NL(SolveOptions.Tolerance)
          Print #4, "time_limit " & str(SolveOptions.MaxTime)

ExitSub:
          Close #4
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "OutputOptionsFile") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

' Adds a new line to the current string, appending LineText at position 0 and CommentText at position CommentSpacing
Sub AddNewLine(CurText As String, LineText As String, Optional CommentText As String = "")
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          CurText = CurText & LineText
          
          ' Add comment with padding if comment should be included
8044      If WriteComments And CommentText <> "" Then
8048          CurText = CurText & Space(CommentSpacing - Len(LineText)) & "# " & CommentText
8049      End If
          
8050      CurText = CurText & vbNewLine

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "AddNewLine") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

' Removes a "\n" character from the end of a string
Private Function StripTrailingNewline(Block As String) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          If right(Block, Len(vbNewLine)) = vbNewLine Then
              Block = left(Block, Len(Block) - Len(vbNewLine))
          End If
          StripTrailingNewline = Block

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "StripTrailingNewline") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Private Function ConvertFormulaToExpressionTree(strFormula As String) As ExpressionTree
          ' Converts a string formula to a complete expression tree
          ' Uses the Shunting Yard algorithm (adapted to produce an expression tree) which takes O(n) time
          ' https://en.wikipedia.org/wiki/Shunting-yard_algorithm
          ' For details on modifications to algorithm
          ' http://wcipeg.com/wiki/Shunting_yard_algorithm#Conversion_into_syntax_tree

          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          Dim tksFormula As Tokens
8056      Set tksFormula = ParseFormula("=" + strFormula)
          
          Dim Operands As New ExpressionTreeStack, Operators As New StringStack, ArgCounts As New OperatorArgCountStack
          
          Dim i As Long, tkn As Token, tknOld As String, Tree As ExpressionTree
8057      For i = 1 To tksFormula.Count
8058          Set tkn = tksFormula.Item(i)
              
8059          Select Case tkn.TokenType
              ' If the token is a number or variable, then add it to the operands stack as a new tree.
              Case TokenType.Number
                  ' Might be a negative number, if so we need to parse out the neg operator
8060              Set Tree = CreateTree(tkn.Text, ExpressionTreeNumber)
8061              If left(Tree.NodeText, 1) = "-" Then
8062                  AddNegToTree Tree
8063              End If
8064              Operands.Push Tree
                      
              Case TokenType.Text
                  Operands.Push CreateTree(tkn.Text, ExpressionTreeString)
8065          Case TokenType.Reference
                  ' TODO this is a hacky way of distinguishing strings eg "A1" from model variables "Sheet_A1"
                  ' Obviously it will fail if the string has an underscore.
                  If InStr(tkn.Text, "_") Then
8066                  Operands.Push CreateTree(tkn.Text, ExpressionTreeVariable)
                  Else
                      Operands.Push CreateTree(tkn.Text, ExpressionTreeString)
                  End If
                  
              ' If the token is a function token, then push it onto the operators stack along with a left parenthesis (tokeniser strips the parenthesis).
8067          Case TokenType.FunctionOpen
8068              Operators.Push ConvertExcelFunctionToNL(tkn.Text)
8069              Operators.Push "("
                  
                  ' Start a new argument count
8070              ArgCounts.PushNewCount
                  
              ' If the token is a function argument separator (e.g., a comma)
8071          Case TokenType.ParameterSeparator
                  ' Until the token at the top of the operator stack is a left parenthesis, pop operators off the stack onto the operands stack as a new tree.
                  ' If no left parentheses are encountered, either the separator was misplaced or parentheses were mismatched.
8072              Do While Operators.Peek() <> "("
8073                  PopOperator Operators, Operands
                      
                      ' If the operator stack runs out without finding a left parenthesis, then there are mismatched parentheses
8074                  If Operators.Count = 0 Then
8075                      Err.Raise OpenSolver_BuildError, "Mismatched parentheses"
8076                  End If
8077              Loop
                  ' Increase arg count for the new parameter
8078              ArgCounts.Increase
              
              ' If the token is an operator
8079          Case TokenType.ArithmeticOperator, TokenType.UnaryOperator, TokenType.ComparisonOperator
                  ' The only unary operator '-' is "neg", which is different to "minus"
8080              If tkn.TokenType = TokenType.UnaryOperator Then
8081                  tkn.Text = "neg"
8082              Else
8083                  tkn.Text = ConvertExcelFunctionToNL(tkn.Text)
8084              End If
                  ' While there is an operator token at the top of the operator stack
8085              Do While Operators.Count > 0
8086                  tknOld = Operators.Peek()
                      ' If either tkn is left-associative and its precedence is less than or equal to that of tknOld
                      ' or tkn has precedence less than that of tknOld
8087                  If CheckPrecedence(tkn.Text, tknOld) Then
                          ' Pop tknOld off the operator stack onto the operand stack as a new tree
8088                      PopOperator Operators, Operands
8089                  Else
8090                      Exit Do
8091                  End If
8092              Loop
                  ' Push operator onto the operator stack
8093              Operators.Push tkn.Text
                  
              ' If the token is a left parenthesis, then push it onto the operator stack
8094          Case TokenType.SubExpressionOpen
8095              Operators.Push tkn.Text
                  
              ' If the token is a right parenthesis
8096          Case TokenType.SubExpressionClose, TokenType.FunctionClose
                  ' Until the token at the top of the operator stack is not a left parenthesis, pop operators off the stack onto the operand stack as a new tree.
8097              Do While Operators.Peek <> "("
8098                  PopOperator Operators, Operands
                      ' If the operator stack runs out without finding a left parenthesis, then there are mismatched parentheses
8099                  If Operators.Count = 0 Then
8100                      Err.Raise OpenSolver_BuildError, "Mismatched parentheses"
8101                  End If
8102              Loop
                  ' Pop the left parenthesis from the operator stack, but not onto the operand stack
8103              Operators.Pop
                  ' If the token at the top of the stack is a function token, pop it onto the operand stack as a new tree
8104              If Operators.Count > 0 Then
8105                  If IsFunctionOperator(Operators.Peek()) Then
8106                      PopOperator Operators, Operands, ArgCounts.PopCount
8107                  End If
8108              End If
8109          End Select
8110      Next i
          
          ' While there are still tokens in the operator stack
8111      Do While Operators.Count > 0
              ' If the token on the top of the operator stack is a parenthesis, then there are mismatched parentheses
8112          If Operators.Peek = "(" Then
8113              Err.Raise OpenSolver_BuildError, "Mismatched parentheses"
8114          End If
              ' Pop the operator onto the operand stack as a new tree
8115          PopOperator Operators, Operands
8116      Loop
          
          ' We are left with a single tree in the operand stack - this is the complete expression tree
8117      Set ConvertFormulaToExpressionTree = Operands.Pop

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "ConvertFormulaToExpressionTree") Then Resume
          RaiseError = True
          GoTo ExitFunction
          
End Function

' Creates an ExpressionTree and initialises NodeText and NodeType
Public Function CreateTree(NodeText As String, NodeType As Long) As ExpressionTree
          Dim obj As ExpressionTree

          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

8120      Set obj = New ExpressionTree
8121      obj.NodeText = NodeText
8122      obj.NodeType = NodeType

8123      Set CreateTree = obj

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "CreateTree") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Function IsNAry(FunctionName As String) As Boolean
8124      Select Case FunctionName
          Case "min", "max", "sum", "count", "numberof", "numberofs", "and_n", "or_n", "alldiff"
8125          IsNAry = True
8126      Case Else
8127          IsNAry = False
8128      End Select
End Function

' Determines the number of operands expected by a .nl operator
Private Function NumberOfOperands(FunctionName As String, Optional ArgCount As Long = 0) As Long
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

8129      Select Case FunctionName
          Case "floor", "ceil", "abs", "neg", "not", "tanh", "tan", "sqrt", "sinh", "sin", "log10", "log", "exp", "cosh", "cos", "atanh", "atan", "asinh", "asin", "acosh", "acos"
8130          NumberOfOperands = 1
8131      Case "plus", "minus", "mult", "div", "rem", "pow", "less", "or", "and", "lt", "le", "eq", "ge", "gt", "ne", "atan2", "intdiv", "precision", "round", "trunc", "iff"
8132          NumberOfOperands = 2
8133      Case "min", "max", "sum", "count", "numberof", "numberofs", "and_n", "or_n", "alldiff"
              'n-ary operator, read number of args from the arg counter
8134          NumberOfOperands = ArgCount
8135      Case "if", "ifs", "implies"
8136          NumberOfOperands = 3
8137      Case Else
8138          Err.Raise OpenSolver_BuildError, "Building expression tree", "Unknown function " & FunctionName & vbCrLf & "Please let us know about this at opensolver.org so we can fix it."
8139      End Select

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "NumberOfOperands") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

' Converts common Excel functinos to .nl operators
Private Function ConvertExcelFunctionToNL(FunctionName As String) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

8140      FunctionName = LCase(FunctionName)
8141      Select Case FunctionName
          Case "ln"
8142          FunctionName = "log"
8143      Case "+"
8144          FunctionName = "plus"
8145      Case "-"
8146          FunctionName = "minus"
8147      Case "*"
8148          FunctionName = "mult"
8149      Case "/"
8150          FunctionName = "div"
8151      Case "mod"
8152          FunctionName = "rem"
8153      Case "^"
8154          FunctionName = "pow"
8155      Case "<"
8156          FunctionName = "lt"
8157      Case "<="
8158          FunctionName = "le"
8159      Case "="
8160          FunctionName = "eq"
8161      Case ">="
8162          FunctionName = "ge"
8163      Case ">"
8164          FunctionName = "gt"
8165      Case "<>"
8166          FunctionName = "ne"
8167      Case "quotient"
8168          FunctionName = "intdiv"
8169      Case "and"
8170          FunctionName = "and_n"
8171      Case "or"
8172          FunctionName = "or_n"
              
8173      Case "log", "ceiling", "floor", "power"
8174          Err.Raise OpenSolver_BuildError, "Building expression tree", "Not implemented yet: " & FunctionName & vbCrLf & "Please let us know about this at opensolver.org so we can fix it."
8175      End Select
8176      ConvertExcelFunctionToNL = FunctionName

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "ConvertExcelFunctionToNL") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

' Determines the precedence of arithmetic operators
Private Function Precedence(tkn As String) As Long
8177      Select Case tkn
          Case "eq", "ne", "gt", "ge", "lt", "le"
              Precedence = 1
          Case "plus", "minus"
8178          Precedence = 2
8179      Case "mult", "div"
8180          Precedence = 3
8181      Case "pow"
8182          Precedence = 4
8183      Case "neg"
8184          Precedence = 5
8185      Case Else
8186          Precedence = -1
8187      End Select
End Function

' Checks the precedence of two operators to determine if the current operator on the stack should be popped
Private Function CheckPrecedence(tkn1 As String, tkn2 As String) As Boolean
          ' Either tkn1 is left-associative and its precedence is less than or equal to that of tkn2
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

8188      If OperatorIsLeftAssociative(tkn1) Then
8189          If Precedence(tkn1) <= Precedence(tkn2) Then
8190              CheckPrecedence = True
8191          Else
8192              CheckPrecedence = False
8193          End If
          ' Or tkn1 has precedence less than that of tkn2
8194      Else
8195          If Precedence(tkn1) < Precedence(tkn2) Then
8196              CheckPrecedence = True
8197          Else
8198              CheckPrecedence = False
8199          End If
8200      End If

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "CheckPrecedence") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

' Determines the left-associativity of arithmetic operators
Private Function OperatorIsLeftAssociative(tkn As String) As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

8201      Select Case tkn
          Case "plus", "minus", "mult", "div", "eq", "ne", "gt", "ge", "lt", "le"
8202          OperatorIsLeftAssociative = True
8203      Case "pow"
8204          OperatorIsLeftAssociative = False
8205      Case Else
8206          Err.Raise OpenSolver_BuildError, "Parsing cleaned formula", "Unknown associativity: " & tkn & vbNewLine & "Please let us know about this at opensolver.org so we can fix it."
8207      End Select

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "OperatorIsLeftAssociative") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

' Pops an operator from the operator stack along with the corresponding number of operands.
Private Sub PopOperator(Operators As StringStack, Operands As ExpressionTreeStack, Optional ArgCount As Long = 0)
          ' Pop the operator and create a new ExpressionTree
          Dim Operator As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

8208      Operator = Operators.Pop()
          Dim NewTree As ExpressionTree
8209      Set NewTree = CreateTree(Operator, ExpressionTreeOperator)
          
          ' Pop the required number of operands from the operand stack and set as children of the new operator tree
          Dim NumToPop As Long, i As Long, Tree As ExpressionTree
8210      NumToPop = NumberOfOperands(Operator, ArgCount)
8211      For i = NumToPop To 1 Step -1
8212          Set Tree = Operands.Pop
8213          NewTree.SetChild i, Tree
8214      Next i
          
          ' Add the new tree to the operands stack
8215      Operands.Push NewTree

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "PopOperator") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

' Check whether a token on the operator stack is a function operator (vs. an arithmetic operator)
Private Function IsFunctionOperator(tkn As String) As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

8216      Select Case tkn
          Case "plus", "minus", "mult", "div", "pow", "neg", "("
8217          IsFunctionOperator = False
8218      Case Else
8219          IsFunctionOperator = True
8220      End Select

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "IsFunctionOperator") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

' Negates a tree by adding a 'neg' node to the root
Private Sub AddNegToTree(Tree As ExpressionTree)
          Dim NewTree As ExpressionTree
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

8221      Set NewTree = CreateTree("neg", ExpressionTreeOperator)
          
8222      Tree.NodeText = Mid(Tree.NodeText, 2)
8223      NewTree.SetChild 1, Tree
          
8224      Set Tree = NewTree

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "AddNegToTree") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

' Formats an expression tree node's text as .nl output
Function FormatNL(NodeText As String, NodeType As ExpressionTreeNodeType) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

8225      Select Case NodeType
          Case ExpressionTreeVariable
8226          FormatNL = "v" & VariableMap(NodeText)
8227      Case ExpressionTreeNumber
8228          FormatNL = "n" & StrEx_NL(Val(NodeText))
8229      Case ExpressionTreeOperator
8230          FormatNL = "o" & CStr(ConvertOperatorToNLCode(NodeText))
8231      End Select

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "FormatNL") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

' Converts an operator string to .nl code
Private Function ConvertOperatorToNLCode(FunctionName As String) As Long
8232      Select Case FunctionName
          Case "plus"
8233          ConvertOperatorToNLCode = 0
8234      Case "minus"
8235          ConvertOperatorToNLCode = 1
8236      Case "mult"
8237          ConvertOperatorToNLCode = 2
8238      Case "div"
8239          ConvertOperatorToNLCode = 3
8240      Case "rem"
8241          ConvertOperatorToNLCode = 4
8242      Case "pow"
8243          ConvertOperatorToNLCode = 5
8244      Case "less"
8245          ConvertOperatorToNLCode = 6
8246      Case "min"
8247          ConvertOperatorToNLCode = 11
8248      Case "max"
8249          ConvertOperatorToNLCode = 12
8250      Case "floor"
8251          ConvertOperatorToNLCode = 13
8252      Case "ceil"
8253          ConvertOperatorToNLCode = 14
8254      Case "abs"
8255          ConvertOperatorToNLCode = 15
8256      Case "neg"
8257          ConvertOperatorToNLCode = 16
8258      Case "or"
8259          ConvertOperatorToNLCode = 20
8260      Case "and"
8261          ConvertOperatorToNLCode = 21
8262      Case "lt"
8263          ConvertOperatorToNLCode = 22
8264      Case "le"
8265          ConvertOperatorToNLCode = 23
8266      Case "eq"
8267          ConvertOperatorToNLCode = 24
8268      Case "ge"
8269          ConvertOperatorToNLCode = 28
8270      Case "gt"
8271          ConvertOperatorToNLCode = 29
8272      Case "ne"
8273          ConvertOperatorToNLCode = 30
8274      Case "if"
8275          ConvertOperatorToNLCode = 35
8276      Case "not"
8277          ConvertOperatorToNLCode = 34
8278      Case "tanh"
8279          ConvertOperatorToNLCode = 37
8280      Case "tan"
8281          ConvertOperatorToNLCode = 38
8282      Case "sqrt"
8283          ConvertOperatorToNLCode = 39
8284      Case "sinh"
8285          ConvertOperatorToNLCode = 40
8286      Case "sin"
8287          ConvertOperatorToNLCode = 41
8288      Case "log10"
8289          ConvertOperatorToNLCode = 42
8290      Case "log"
8291          ConvertOperatorToNLCode = 43
8292      Case "exp"
8293          ConvertOperatorToNLCode = 44
8294      Case "cosh"
8295          ConvertOperatorToNLCode = 45
8296      Case "cos"
8297          ConvertOperatorToNLCode = 46
8298      Case "atanh"
8299          ConvertOperatorToNLCode = 47
8300      Case "atan2"
8301          ConvertOperatorToNLCode = 48
8302      Case "atan"
8303          ConvertOperatorToNLCode = 49
8304      Case "asinh"
8305          ConvertOperatorToNLCode = 50
8306      Case "asin"
8307          ConvertOperatorToNLCode = 51
8308      Case "acosh"
8309          ConvertOperatorToNLCode = 52
8310      Case "acos"
8311          ConvertOperatorToNLCode = 53
8312      Case "sum"
8313          ConvertOperatorToNLCode = 54
8314      Case "intdiv"
8315          ConvertOperatorToNLCode = 55
8316      Case "precision"
8317          ConvertOperatorToNLCode = 56
8318      Case "round"
8319          ConvertOperatorToNLCode = 57
8320      Case "trunc"
8321          ConvertOperatorToNLCode = 58
8322      Case "count"
8323          ConvertOperatorToNLCode = 59
8324      Case "numberof"
8325          ConvertOperatorToNLCode = 60
8326      Case "numberofs"
8327          ConvertOperatorToNLCode = 61
8328      Case "ifs"
8329          ConvertOperatorToNLCode = 65
8330      Case "and_n"
8331          ConvertOperatorToNLCode = 70
8332      Case "or_n"
8333          ConvertOperatorToNLCode = 71
8334      Case "implies"
8335          ConvertOperatorToNLCode = 72
8336      Case "iff"
8337          ConvertOperatorToNLCode = 73
8338      Case "alldiff"
8339          ConvertOperatorToNLCode = 74
8340      End Select
End Function

' Converts an objective sense to .nl code
Private Function ConvertObjectiveSenseToNL(ObjectiveSense As ObjectiveSenseType) As Long
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

8341      Select Case ObjectiveSense
          Case ObjectiveSenseType.MaximiseObjective
8342          ConvertObjectiveSenseToNL = 1
8343      Case ObjectiveSenseType.MinimiseObjective
8344          ConvertObjectiveSenseToNL = 0
8345      Case Else
8346          Err.Raise OpenSolver_SolveError, Description:="Objective sense not supported: " & ObjectiveSense
8347      End Select

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "ConvertObjectiveSenseToNL") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

' Converts RelationConsts enum to .nl code.
Private Sub ConvertConstraintToNL(Relation As RelationConsts, BoundType As Long, Comment As String)
8348      Select Case Relation
              Case RelationConsts.RelationLE ' Upper Bound on LHS
8349              BoundType = 1
8350              Comment = " <= "
8351          Case RelationConsts.RelationEQ ' Equality
8352              BoundType = 4
8353              Comment = " == "
8354          Case RelationConsts.RelationGE ' Upper Bound on RHS
8355              BoundType = 2
8356              Comment = " >= "
8357      End Select
End Sub

Function ReadModel_NL(SolutionFilePathName As String, errorString As String, s As COpenSolverParsed) As Boolean
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    ReadModel_NL = False
    
    Dim Line As String
    Dim solutionExpected As Boolean
    solutionExpected = True
    
    If Not FileOrDirExists(SolutionFilePathName) Then
        solutionExpected = False
        If Not TryParseLogs(s) Then
            errorString = "The solver did not create a solution file. No new solution is available." & vbCrLf & vbCrLf & _
                          "This can happen when the initial conditions are invalid. " & _
                          "Check the log file for more information."
            GoTo ExitFunction
        End If
    Else
        Open SolutionFilePathName For Input As #1
        Line Input #1, Line ' Skip empty line at start of file
        Line Input #1, Line ' Get line with status code

        'Get the returned status code from solver.
        If InStrText(Line, "optimal") Then
            s.SolveStatus = OpenSolverResult.Optimal
            s.SolveStatusString = "Optimal"
        ElseIf InStrText(Line, "infeasible") Then
            s.SolveStatus = OpenSolverResult.Infeasible
            s.SolveStatusString = "No Feasible Solution"
            solutionExpected = False
        ElseIf InStrText(Line, "unbounded") Then
            s.SolveStatus = OpenSolverResult.Unbounded
            s.SolveStatusString = "No Solution Found (Unbounded)"
            solutionExpected = False
        ElseIf InStrText(Line, "interrupted on limit") Then
            s.SolveStatus = OpenSolverResult.LimitedSubOptimal
            s.SolveStatusString = "Stopped on User Limit (Time/Iterations)"
            ' See if we can find out which limit was hit from the log file
            TryParseLogs s
        ElseIf InStrText(Line, "interrupted by user") Then
            s.SolveStatus = OpenSolverResult.AbortedThruUserAction
            s.SolveStatusString = "Stopped on Ctrl-C"
        Else
            If Not TryParseLogs(s) Then
                errorString = "The response from the " & SolverName(m.Solver) & " solver is not recognised. The response was: " & vbCrLf & vbCrLf & _
                              Line & vbCrLf & vbCrLf & _
                              "The " & SolverName(m.Solver) & " command line can be found at:" & vbCrLf & _
                              ScriptFilePath(m.Solver)
                s.SolveStatus = OpenSolverResult.ErrorOccurred
                GoTo ExitFunction
            End If
            solutionExpected = False
        End If
    End If
    
    If solutionExpected Then
        UpdateStatusBar "OpenSolver: Loading Solution... " & s.SolveStatusString, True
        
        Line Input #1, Line ' Throw away blank line
        Line Input #1, Line ' Throw away "Options"
        
        Dim i As Long
        For i = 1 To 8
            Line Input #1, Line ' Skip all options lines
        Next i
        
        ' Note that the variable values are written to file in .nl format
        ' We need to read in the values and the extract the correct values for the adjustable cells
        
        ' Read in all variable values
        Dim VariableValues As New Collection
        While Not EOF(1)
            Line Input #1, Line
            VariableValues.Add Val(Line)
        Wend
        
        ' Loop through variable cells and find the corresponding value from VariableValues
        i = 1
        Dim c As Range, VariableIndex As Long
        For Each c In m.AdjustableCells
            ' Extract the correct variable value
            VariableIndex = GetVariableNLIndex(i) + 1
            
            ' Need to make sure number is in US locale when Value2 is set
            Range(c.Address).Value2 = ConvertFromCurrentLocale(VariableValues(VariableIndex))
            i = i + 1
        Next c
        s.SolutionWasLoaded = True
    End If

    ReadModel_NL = True

ExitFunction:
    Application.StatusBar = False
    Close #1
    Close #2
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("ParsedSolverNL", "ReadModel_NL") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function

Private Function TryParseLogs(s As COpenSolverParsed) As Boolean
      ' We examine the log file if it exists to try to find more info about the solve
          
          ' Check if log exists
          Dim logFile As String
8477      logFile = GetTempFilePath("log1.tmp")
          
          On Error GoTo ErrorHandler
8478      If Not FileOrDirExists(logFile) Then
8479          TryParseLogs = False
8480          GoTo ExitFunction
8481      End If
          
          Dim message As String
8483      Open logFile For Input As #3
8484      message = Input$(LOF(3), 3)
8485      Close #3
          
          ' We need to check > 0 explicitly, as the expression doesn't work without it
8486      If Not InStrText(left(message, 7), SolverName(m.Solver)) > 0 Then
             ' Not dealing with the correct solver log, abort
8487          TryParseLogs = False
8488          GoTo ExitFunction
8489      End If
          
          ' Scan for information
          
          ' 1 - scan for time limit
          If InStrText(message, "exiting on maximum time") Then
              s.SolveStatus = OpenSolverResult.LimitedSubOptimal
              s.SolveStatusString = "Stopped on Time Limit"
              TryParseLogs = True
              GoTo ExitFunction
          End If
          ' 2 - scan for iteration limit
          If InStrText(message, "exiting on maximum number of iterations") Then
              s.SolveStatus = OpenSolverResult.LimitedSubOptimal
              s.SolveStatusString = "Stopped on Iteration Limit"
              TryParseLogs = True
              GoTo ExitFunction
          End If
          ' 3 - scan for infeasible
8490      If InStrText(message, "infeasible") Then
8491          s.SolveStatus = OpenSolverResult.Infeasible
8492          s.SolveStatusString = "No Feasible Solution"
8493          TryParseLogs = True
8494          GoTo ExitFunction
8495      End If

ExitFunction:
8496      Close #3
          Exit Function

ErrorHandler:
          If Not ReportError("ParsedSolverNL", "TryParseLogs") Then Resume
          TryParseLogs = False  ' Mark read as fail
          GoTo ExitFunction
End Function

Function CreateSolveScript_NL(ModelFilePathName As String, SolveOptions As SolveOptionsType) As String
    ' Create a script to cd to temp and run "/path/to/solver /path/to/<ModelFilePathName>"
    
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    Dim SolverString As String, CommandLineRunString As String
    SolverAvailable m.Solver, SolverString
    SolverString = MakePathSafe(SolverString)
    
    CommandLineRunString = MakePathSafe(ModelFilePathName)
       
    Dim scriptFile As String, scriptFileContents As String
    scriptFile = ScriptFilePath(m.Solver)
    
    scriptFileContents = "cd " & MakePathSafe(GetTempFolder()) & " && " & _
                         SolverString & " " & CommandLineRunString
    CreateScriptFile scriptFile, scriptFileContents
       
    CreateSolveScript_NL = scriptFile
    
    ' Create the options file in the temp folder
    OutputOptionsFile OptionsFilePath(m.Solver), SolveOptions

ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("ParsedSolverNL", "CreateSolveScript_NL") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function

