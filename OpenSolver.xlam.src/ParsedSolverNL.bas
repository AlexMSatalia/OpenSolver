Attribute VB_Name = "ParsedSolverNL"
Option Explicit

' This module is for writing .nl files that describe the model and solving these
' For more info on .nl file format see:
' http://citeseerx.ist.psu.edu/viewdoc/summary?doi=10.1.1.60.9659
' http://www.ampl.com/REFS/hooking2.pdf

Dim m As CModelParsed

Dim problem_name As String

Public WriteComments As Boolean     ' Whether .nl file should include comments
Public CommentIndent As Long     ' Tracks the level of indenting in comments on nl output
Public Const CommentSpacing = 24    ' The column number at which nl comments begin

Dim errorPrefix As String

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

Dim LinearConstraints As Collection     ' Collection containing the LinearConstraintNLs for each constraint stored in parsed constraint order
Dim LinearConstants As Collection       ' Collection containing the constant (a Double) for each constraint stored in parsed constraint order
Dim LinearObjectives As Collection      ' Collection containing the LinearConstraintNLs for each objective stored in parsed objective order

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
Public Property Get GetVariableNLIndex(index As Long) As Long
7458      GetVariableNLIndex = VariableCollectionIndexToNLIndex(index)
End Property

' Creates .nl file and solves model
Function SolveModelParsed_NL(ModelFilePathName As String, model As CModelParsed, s As COpenSolverParsed, SolveOptions As SolveOptionsType, Optional ShouldWriteComments As Boolean = True)
7459      Set m = model
          
7460      WriteComments = ShouldWriteComments
          
7461      errorPrefix = "Constructing .nl file"
          
          ' =============================================================
          ' Process model for .nl output
          ' All Module-level variables required for .nl output should be set in this step
          ' No modification to these variables should be done while writing the .nl file
          ' =============================================================
          
7462      InitialiseModelStats
          
7463      On Error GoTo ErrHandler
7464      If n_obj = 0 And n_con = 0 Then
7465          Err.Raise OpenSolver_BuildError, errorPrefix, "The model has no constraints that depend on the adjustable cells, and has no objective. There is nothing for the solver to do."
7466      End If
          
7467      CreateVariableIndex
          
7468      ProcessFormulae
7469      ProcessObjective
          
7470      MakeVariableMap
7471      MakeConstraintMap
          
          ' =============================================================
          ' Write output files
          ' =============================================================
          
          ' Create supplementary outputs
7472      OutputColFile
7473      OutputRowFile
          
          ' Write .nl file
7474      Open ModelFilePathName For Output As #1
          
          ' Write header
7475      Print #1, MakeHeader()
          ' Write C blocks
7476      If n_con > 0 Then
7477          Print #1, MakeCBlocks()
7478      End If
          
7479      If n_obj > 0 Then
              ' Write O block
7480          Print #1, MakeOBlocks()
7481      End If
          
'          ' Write d block
'7482      If n_con > 0 Then
'7483          Print #1, MakeDBlock()
'7484      End If
          
          ' Write x block
7485      Print #1, MakeXBlock()
          
          ' Write r block
7486      If n_con > 0 Then
7487          Print #1, MakeRBlock()
7488      End If
          
          ' Write b block
7489      Print #1, MakeBBlock()
          
          ' Write k block
7490      If n_con > 0 Then
7491          Print #1, MakeKBlock()
7492      End If
          
          ' Write J block
7493      If n_con > 0 Then
7494          Print #1, MakeJBlocks()
7495      End If
          
7496      If n_obj > 0 Then
              ' Write G block
7497          Print #1, MakeGBlocks()
7498      End If
          
7499      Close #1
          
          ' =============================================================
          ' Solve model using chosen solver
          ' =============================================================
          
7500      errorPrefix = "Solving .nl model file"
7501      Application.StatusBar = "OpenSolver: " & errorPrefix
          
          Dim SolutionFilePathName As String
7502      SolutionFilePathName = SolutionFilePath(m.Solver)
          
7503      DeleteFileAndVerify SolutionFilePathName, errorPrefix, "Unable to delete solution file : " & SolutionFilePathName
          
          Dim ExternalSolverPathName As String
7504      ExternalSolverPathName = CreateSolveScriptParsed(m.Solver, ModelFilePathName, SolveOptions)
                   
          Dim logCommand As String, logFileName As String
7505      logFileName = "log1.tmp"
7506      logCommand = MakePathSafe(GetTempFilePath(logFileName))
                        
          Dim ExecutionCompleted As Boolean
7507      ExternalSolverPathName = MakePathSafe(ExternalSolverPathName)
                    
          Dim exeResult As Long, userCancelled As Boolean
7508      ExecutionCompleted = RunExternalCommand(ExternalSolverPathName, logCommand, IIf(s.GetShowIterationResults, SW_SHOWNORMAL, SW_HIDE), True, userCancelled, exeResult) ' Run solver, waiting for completion
7509      If userCancelled Then
              ' User pressed escape. Dialogs have already been shown. Exit with a 'cancelled' error
7510          On Error GoTo ErrHandler
7511          Err.Raise Number:=OpenSolver_UserCancelledError, Source:="Solving NL model", Description:="The solving process was cancelled by the user."
7512      End If
7513      If exeResult <> 0 Then
              ' User pressed escape. Dialogs have already been shown. Exit with a 'cancelled' error
7514          On Error GoTo ErrHandler
7515          Err.Raise Number:=OpenSolver_SolveError, Source:="Solving NL model", Description:="The " & m.Solver & " solver did not complete, but aborted with the error code " & exeResult & "." & vbCrLf & vbCrLf & "The last log file can be viewed under the OpenSolver menu and may give you more information on what caused this error."
7516      End If
          
          ' =============================================================
          ' Read results from solution file
          ' =============================================================
7517      errorPrefix = "Reading .nl solution"
7518      Application.StatusBar = "OpenSolver: " & errorPrefix

          Dim solutionLoaded As Boolean, errorString As String
7519      solutionLoaded = ReadModelParsed(m.Solver, SolutionFilePathName, errorString, m, s)
7520      On Error GoTo ErrHandler
7521      If errorString <> "" Then
7522          Err.Raise Number:=OpenSolver_SolveError, Source:="Solving NL model", Description:=errorString
7523      ElseIf Not solutionLoaded Then 'read error
7524          SolveModelParsed_NL = False
7525          Exit Function
7526      End If

7527      SolveModelParsed_NL = True
7528      Exit Function
          
exitFunction:
7529      Application.StatusBar = False
7530      Close #1
7531      Exit Function
              
ErrHandler:
          ' We only trap Escape (Err.Number=18) here; all other errors are passed back to the caller.
          ' Save error message
          Dim ErrorNumber As Long, ErrorDescription As String, ErrorSource As String
7532      ErrorNumber = Err.Number
7533      ErrorDescription = Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
7534      ErrorSource = Err.Source

7535      If Err.Number = 18 Then
7536          If MsgBox("You have pressed the Escape key. Do you wish to cancel?", _
                         vbCritical + vbYesNo + vbDefaultButton1, _
                         "OpenSolver: User Interrupt Occured...") = vbNo Then
7537              Resume 'continue on from where error occured
7538          Else
                  ' Raise a "user cancelled" error. We cannot use Raise, as that exits immediately without going thru our code below
7539              ErrorNumber = OpenSolver_UserCancelledError
7540              ErrorSource = "Parsing formulae"
7541              ErrorDescription = "Model building cancelled by user."
7542          End If
7543      End If
          
ErrorExit:
          ' Exit, raising an error; none of the following change the Err.Number etc, but we saved them above just in case...
7544      Application.StatusBar = False
7545      Close #1
7546      Err.Raise ErrorNumber, ErrorSource, ErrorDescription
End Function

Sub InitialiseModelStats()
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
7554          If i Mod 100 = 1 Then
7555              Application.StatusBar = "OpenSolver: Creating .nl file. Counting constraints: " & i & "/" & numActualCons & ". "
7556          End If
7557          DoEvents
              
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
End Sub

' Creates map from variable name (e.g. Test1_D4) to parsed variable index (1 to n_var)
Sub CreateVariableIndex()
7596      Set VariableIndex = New Collection
7597      Set InitialVariableValues = New Collection
          Dim c As Range, cellName As String, i As Long
          
          
          ' First read in actual vars
7598      i = 1
7599      For Each c In m.AdjustableCells
7600          If i Mod 100 = 1 Then
7601              Application.StatusBar = "OpenSolver: Creating .nl file. Counting variables: " & i & "/" & numActualVars & ". "
7602          End If
7603          DoEvents
              
7604          cellName = ConvertCellToStandardName(c)
              
              ' Update variable maps
7605          VariableIndex.Add i, cellName
              
              ' Update initial values
7606          InitialVariableValues.Add CDbl(c)
              
7607          i = i + 1
7608      Next c
          
          ' Next read in fake formulae vars
7609      For i = 1 To numFakeVars
7610          If i Mod 100 = 1 Then
7611              Application.StatusBar = "OpenSolver: Creating .nl file. Counting formulae variables: " & i & "/" & numFakeVars & ". "
7612          End If
7613          DoEvents
              
7614          cellName = m.Formulae(i).strAddress
              
              ' Update variable maps
7615          VariableIndex.Add i + numActualVars, cellName
7616      Next i
End Sub

' Creates maps from variable name (e.g. Test1_D4) to .nl variable index (0 to n_var - 1) and vice-versa, and
' maps from parsed variable index to .nl variable index and vice-versa
Sub MakeVariableMap()
          ' =============================================
          ' Create index of variable names in parsed variable order
          Dim CellNames As New Collection
          
          ' Actual variables
          Dim c As Range, cellName As String, i As Long
7617      i = 1
7618      For Each c In m.AdjustableCells
7619          If i Mod 100 = 1 Then
7620              Application.StatusBar = "OpenSolver: Creating .nl file. Classifying variables: " & i & "/" & numActualVars & ". "
7621          End If
7622          DoEvents
              
7623          i = i + 1
7624          cellName = ConvertCellToStandardName(c)
7625          CellNames.Add cellName
7626      Next c
          
          ' Formulae variables
7627      For i = 1 To m.Formulae.Count
7628          If i Mod 100 = 1 Then
7629              Application.StatusBar = "OpenSolver: Creating .nl file. Classifying formulae variables: " & i & "/" & numFakeVars & ". "
7630          End If
7631          DoEvents
              
7632          cellName = m.Formulae(i).strAddress
7633          CellNames.Add cellName
7634      Next i
          
          '==============================================
          ' Get binary and integer variables from model
          
          ' Get integer variables
          Dim IntegerVars() As Boolean
7635      ReDim IntegerVars(n_var)
7636      If Not m.IntegerCells Is Nothing Then
7637          For Each c In m.IntegerCells
7638              Application.StatusBar = "OpenSolver: Creating .nl file. Finding integer variables"
7639              DoEvents
              
7640              cellName = ConvertCellToStandardName(c)
7641              IntegerVars(VariableIndex(cellName)) = True
7642          Next c
7643      End If
          
          ' Get binary variables
7644      ReDim BinaryVars(n_var)
7645      If Not m.BinaryCells Is Nothing Then
7646          For Each c In m.BinaryCells
7647              Application.StatusBar = "OpenSolver: Creating .nl file. Finding binary variables"
7648              DoEvents
                  
7649              cellName = ConvertCellToStandardName(c)
7650              BinaryVars(VariableIndex(cellName)) = True
                  ' Reset integer state for this variable - binary trumps integer
7651              IntegerVars(VariableIndex(cellName)) = False
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
7655          If i Mod 100 = 1 Then
7656              Application.StatusBar = "OpenSolver: Creating .nl file. Sorting variables: " & i & "/" & n_var & ". "
7657          End If
7658          DoEvents
              
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
          
          Dim index As Long, var As Long
7679      index = 0
          
          ' We loop through the variables and arrange them in the required order:
          '     1st - non-linear continuous
          '     2nd - non-linear integer
          '     3rd - linear arcs (N/A)
          '     4th - other linear
          '     5th - binary
          '     6th - other integer
          
          ' Non-linear continuous
7680      For i = 1 To NonLinearContinuous.Count
7681          If i Mod 100 = 1 Then
7682              Application.StatusBar = "OpenSolver: Creating .nl file. Outputting non-linear continuous vars"
7683          End If
7684          DoEvents
              
7685          var = NonLinearContinuous(i)
7686          AddVariable CellNames(var), index, var
7687      Next i
          
          ' Non-linear integer
7688      For i = 1 To NonLinearInteger.Count
7689          If i Mod 100 = 1 Then
7690              Application.StatusBar = "OpenSolver: Creating .nl file. Outputting non-linear integer vars"
7691          End If
7692          DoEvents
              
7693          var = NonLinearInteger(i)
7694          AddVariable CellNames(var), index, var
7695      Next i
          
          ' Linear continuous
7696      For i = 1 To LinearContinuous.Count
7697          If i Mod 100 = 1 Then
7698              Application.StatusBar = "OpenSolver: Creating .nl file. Outputting linear continuous vars"
7699          End If
7700          DoEvents
              
7701          var = LinearContinuous(i)
7702          AddVariable CellNames(var), index, var
7703      Next i
          
          ' Linear binary
7704      For i = 1 To LinearBinary.Count
7705          If i Mod 100 = 1 Then
7706              Application.StatusBar = "OpenSolver: Creating .nl file. Outputting linear binary vars"
7707          End If
7708          DoEvents
              
7709          var = LinearBinary(i)
7710          AddVariable CellNames(var), index, var
7711      Next i
          
          ' Linear integer
7712      For i = 1 To LinearInteger.Count
7713          If i Mod 100 = 1 Then
7714              Application.StatusBar = "OpenSolver: Creating .nl file. Outputting linear integer vars"
7715          End If
7716          DoEvents
              
7717          var = LinearInteger(i)
7718          AddVariable CellNames(var), index, var
7719      Next i
          
          ' ==============================================
          ' Update model stats
          
7720      nbv = LinearBinary.Count
7721      niv = LinearInteger.Count
7722      nlvci = NonLinearInteger.Count
End Sub

' Adds a variable to all the variable maps with:
'   variable name:            cellName
'   .nl variable index:       index
'   parsed variable index:    i
Sub AddVariable(cellName As String, index As Long, i As Long)
          ' Update variable maps
7723      VariableMap.Add CStr(index), cellName
7724      VariableMapRev.Add cellName, CStr(index)
7725      VariableNLIndexToCollectionIndex(index) = i
7726      VariableCollectionIndexToNLIndex(i) = index
          
          ' Update max length of variable name
7727      If Len(cellName) > maxcolnamelen_ Then
7728         maxcolnamelen_ = Len(cellName)
7729      End If
          
          ' Increase index for the next variable
7730      index = index + 1
End Sub

' Creates maps from constraint name (e.g. c1_Test1_D4) to .nl constraint index (0 to n_con - 1) and vice-versa, and
' map from .nl constraint index to parsed constraint index
Sub MakeConstraintMap()
          
7731      Set ConstraintMap = New Collection
7732      Set ConstraintMapRev = New Collection
7733      ReDim ConstraintIndexToTreeIndex(n_con)
          
          Dim index As Long, i As Long, cellName As String
7734      index = 0
          
          ' We loop through the constraints and arrange them in the required order:
          '     1st - non-linear
          '     2nd - non-linear network (N/A)
          '     3rd - linear network (N/A)
          '     4th - linear
          
          ' Non-linear constraints
7735      For i = 1 To n_con
7736          If i Mod 100 = 1 Then
7737              Application.StatusBar = "OpenSolver: Creating .nl file. Outputting non-linear constraints" & i & "/" & n_con
7738          End If
7739          DoEvents
              
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
7748              AddConstraint cellName, index, i
7749          End If
7750      Next i
          
          ' Linear constraints
7751      For i = 1 To n_con
7752          If i Mod 100 = 1 Then
7753              Application.StatusBar = "OpenSolver: Creating .nl file. Outputting linear constraints" & i & "/" & n_con
7754          End If
7755          DoEvents
              
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
7764              AddConstraint cellName, index, i
7765          End If
7766      Next i
End Sub

' Adds a constraint to all the constraint maps with:
'   constraint name:          cellName
'   .nl constraint index:     index
'   parsed constraint index:  i
Sub AddConstraint(cellName As String, index As Long, i As Long)
          ' Update constraint maps
7767      ConstraintMap.Add index, cellName
7768      ConstraintMapRev.Add cellName, CStr(index)
7769      ConstraintIndexToTreeIndex(index) = i
          
          ' Update max length
7770      If Len(cellName) > maxrownamelen_ Then
7771         maxrownamelen_ = Len(cellName)
7772      End If
          
          ' Increase index for next constraint
7773      index = index + 1
End Sub

' Processes all constraint formulae into the .nl model formats
Sub ProcessFormulae()
          
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
7782          If i Mod 100 = 1 Then
7783              Application.StatusBar = "OpenSolver: Processing formulae into expression trees... " & i & "/" & n_con & " formulae."
7784          End If
7785          DoEvents
              
7786          ProcessSingleFormula m.RHSKeys(i), m.LHSKeys(i), m.Rels(i)
7787      Next i
          
7788      For i = 1 To numFakeCons
7789          If i Mod 100 = 1 Then
7790              Application.StatusBar = "OpenSolver: Processing formulae into expression trees... " & i + numActualCons & "/" & n_con & " formulae."
7791          End If
7792          DoEvents
              
7793          ProcessSingleFormula m.Formulae(i).strFormulaParsed, m.Formulae(i).strAddress, RelationConsts.RelationEQ
7794      Next i
          
7795      Application.StatusBar = False
End Sub

' Processes a single constraint into .nl format. We require:
'     - a non-linear ExpressionTree for all non-linear parts of the equation
'     - a linear LinearConstraintNL for the linear parts of the equation
'     - a constant Double for the constant part of the equation
' We also use the results of processing to update some of the model statistics
Sub ProcessSingleFormula(RHSExpression As String, LHSVariable As String, Relation As RelationConsts)
          ' Convert the string formula into an ExpressionTree object
          Dim Tree As ExpressionTree
7796      Set Tree = ConvertFormulaToExpressionTree(RHSExpression)
          
          ' The .nl file needs a linear coefficient for every variable in the constraint - non-linear or otherwise
          ' We need a list of all variables in this constraint so that we can know to include them in the linear part of the constraint.
          Dim constraint As New LinearConstraintNL
7797      constraint.Count = n_var
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
7810          constraint.VariablePresent(VariableIndex(LHSVariable)) = True
7811          constraint.Coefficient(VariableIndex(LHSVariable)) = constraint.Coefficient(VariableIndex(LHSVariable)) - 1
7812      End If
          
          ' Keep constant on RHS and take everything else to LHS.
          ' Need to flip all sign of elements in this constraint

          ' Flip all coefficients in linear constraint
7821      constraint.InvertCoefficients

          ' Negate non-linear tree
7822      Set Tree = Tree.Negate
          
          ' Save results of processing
7824      NonLinearConstraintTrees.Add Tree
7825      LinearConstraints.Add constraint
7826      LinearConstants.Add constant
7827      ConstraintRelations.Add Relation
          
          ' Mark any non-linear variables that we haven't seen before by extracting any remaining variables from the non-linear tree.
          ' Any variable still present in the constraint must be part of the non-linear section
          Dim TempConstraint As New LinearConstraintNL
7828      TempConstraint.Count = n_var
7829      Tree.ExtractVariables TempConstraint
7830      For j = 1 To TempConstraint.Count
7831          If TempConstraint.VariablePresent(j) And Not NonLinearVars(j) Then
7832              NonLinearVars(j) = True
7833              nlvc = nlvc + 1
7834          End If
7835      Next j
          
          ' Remove any zero coefficients from the linear constraint if the variable is not in the non-linear tree
7836      For j = 1 To constraint.Count
7837          If Not NonLinearVars(j) And constraint.Coefficient(j) = 0 Then
7838              constraint.VariablePresent(j) = False
7839          End If
7840      Next j
          
          ' Increase count of non-linear constraints if the non-linear tree is non-empty
          ' An empty tree has a single "0" node
7841      If Tree.NodeText <> "0" Then
              Dim ConstraintIndex As Long
7842          ConstraintIndex = NonLinearConstraintTrees.Count
7843          NonLinearConstraints(ConstraintIndex) = True
7844          nlc = nlc + 1
7845      End If
          
          ' Update jacobian counts using the linear variables present
7846      For j = 1 To constraint.Count
              ' The nl documentation says that the jacobian counts relate to "the numbers of nonzeros in the first n var - 1 columns of the Jacobian matrix"
              ' This means we should increases the count for any variable that is present and has a non-zero coefficient
              ' However, .nl files generated by AMPL seem to increase the count for any present variable, even if the coefficient is zero (i.e. non-linear variables)
              ' We adopt the AMPL behaviour as it gives faster solve times (see Test 40). Swap the If conditions below to change this behaviour
              
              'If constraint.VariablePresent(j) And constraint.Coefficient(j) <> 0 Then
7847          If constraint.VariablePresent(j) Then
7848              NonZeroConstraintCount(j) = NonZeroConstraintCount(j) + 1
7849              nzc = nzc + 1
7850          End If
7851      Next j
End Sub

' Process objective function into .nl format. We require the same as for constraints:
'     - a non-linear ExpressionTree for all non-linear parts of the equation
'     - a linear LinearConstraintNL for the linear parts of the equation
'     - a constant Double for the constant part of the equation
Sub ProcessObjective()
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
              Dim Objective As New LinearConstraintNL
7865          Objective.Count = n_var
7866          Objective.VariablePresent(VariableIndex(ObjectiveCells(1))) = True
7867          Objective.Coefficient(VariableIndex(ObjectiveCells(1))) = 1
              
              ' Save results
7868          LinearObjectives.Add Objective
              
              ' Track non-zero jacobian count in objective
7869          nzo = nzo + 1
              
              ' Adjust objective count
7870          n_obj = n_obj + 1
7871      End If
End Sub

' Writes header block for .nl file. This contains the model statistics
Function MakeHeader() As String
          Dim Header As String
7872      Header = ""
          
          'Line #1
7873      AddNewLine Header, "g3 1 1 0", "problem " & problem_name
          'Line #2
7874      AddNewLine Header, " " & n_var & " " & n_con & " " & n_obj & " " & nranges & " " & n_eqn_ & " " & n_lcon, "vars, constraints, objectives, ranges, eqns"
          'Line #3
7875      AddNewLine Header, " " & nlc & " " & nlo, "nonlinear constraints, objectives"
          'Line #4
7876      AddNewLine Header, " " & nlnc & " " & lnc, "network constraints: nonlinear, linear"
          'Line #5
7877      AddNewLine Header, " " & nlvc & " " & nlvo & " " & nlvb, "nonlinear vars in constraints, objectives, both"
          'Line #6
7878      AddNewLine Header, " " & nwv_ & " " & nfunc_ & " " & arith & " " & flags, "linear network variables; functions; arith, flags"
          'Line #7
7879      AddNewLine Header, " " & nbv & " " & niv & " " & nlvbi & " " & nlvci & " " & nlvoi, "discrete variables: binary, integer, nonlinear (b,c,o)"
          'Line #8
7880      AddNewLine Header, " " & nzc & " " & nzo, "nonzeros in Jacobian, gradients"
          'Line #9
7881      AddNewLine Header, " " & maxrownamelen_ & " " & maxcolnamelen_, "max name lengths: constraints, variables"
          'Line #10
7882      AddNewLine Header, " " & comb & " " & comc & " " & como & " " & comc1 & " " & como1, "common exprs: b,c,o,c1,o1"
          
7883      MakeHeader = StripTrailingNewline(Header)
End Function

' Writes C blocks for .nl file. These describe the non-linear parts of each constraint.
Function MakeCBlocks() As String
          Dim Block As String
7884      Block = ""
          
          Dim i As Long
7885      For i = 1 To n_con
7886          If i Mod 100 = 1 Then
7887              Application.StatusBar = "OpenSolver: Creating .nl file. Writing non-linear constraints" & i & "/" & n_con
7888          End If
7889          DoEvents
              
              ' Add block header for the constraint
7890          AddNewLine Block, "C" & i - 1, "CONSTRAINT NON-LINEAR SECTION " + ConstraintMapRev(CStr(i - 1))
              
              ' Add expression tree
7891          CommentIndent = 4
7892          Block = Block + NonLinearConstraintTrees(ConstraintIndexToTreeIndex(i - 1)).ConvertToNL
7893      Next i
          
7894      MakeCBlocks = StripTrailingNewline(Block)
End Function

' Writes O blocks for .nl file. These describe the non-linear parts of each objective.
Function MakeOBlocks() As String
          Dim Block As String
7895      Block = ""
          
          Dim i As Long
7896      For i = 1 To n_obj
              ' Add block header for the objective
7897          AddNewLine Block, "O" & i - 1 & " " & ConvertObjectiveSenseToNL(ObjectiveSenses(i)), "OBJECTIVE NON-LINEAR SECTION " & ObjectiveCells(i)
              
              ' Add expression tree
7898          CommentIndent = 4
7899          Block = Block + NonLinearObjectiveTrees(i).ConvertToNL
7900      Next i
          
7901      MakeOBlocks = StripTrailingNewline(Block)
End Function

' Writes D block for .nl file. This contains the initial guess for dual variables.
' We don't use this, so just set them all to zero
Function MakeDBlock() As String
          Dim Block As String
7902      Block = ""
          
          ' Add block header
7903      AddNewLine Block, "d" & n_con, "INITIAL DUAL GUESS"
          
          ' Set duals to zero for all constraints
          Dim i As Long
7904      For i = 1 To n_con
7905          If i Mod 100 = 1 Then
7906              Application.StatusBar = "OpenSolver: Creating .nl file. Writing initial duals " & i & "/" & n_con
7907          End If
7908          DoEvents
              
7909          AddNewLine Block, i - 1 & " 0", "    " & ConstraintMapRev(CStr(i - 1)) & " = " & 0
7910      Next i
          
7911      MakeDBlock = StripTrailingNewline(Block)
End Function

' Writes X block for .nl file. This contains the initial guess for primal variables
Function MakeXBlock() As String
          Dim Block As String
7912      Block = ""
          
          ' Add block header
7913      AddNewLine Block, "x" & n_var, "INITIAL PRIMAL GUESS"

          ' Loop through the variables in .nl variable order
          Dim i As Long, initial As String, VariableIndex As Long
7914      For i = 1 To n_var
7915          If i Mod 100 = 1 Then
7916              Application.StatusBar = "OpenSolver: Creating .nl file. Writing initial values " & i & "/" & n_var
7917          End If
7918          DoEvents
              
7919          VariableIndex = VariableNLIndexToCollectionIndex(i - 1)
              
              ' Get initial values
7920          If VariableIndex <= numActualVars Then
                  ' Actual variables - use the value in the actual cell
7921              initial = InitialVariableValues(VariableIndex)
7922          Else
                  ' Formulae variables - use the initial value saved in the CFormula instance
7923              initial = CDbl(m.Formulae(VariableIndex - numActualVars).initialValue)
7924          End If
7925          AddNewLine Block, i - 1 & " " & initial, "    " & VariableMapRev(CStr(i - 1)) & " = " & initial
7926      Next i
          
7927      MakeXBlock = StripTrailingNewline(Block)
End Function

' Writes R block for .nl file. This contains the constant values for each constraint and the relation type
Function MakeRBlock() As String
          Dim Block As String
7928      Block = ""
           ' Add block header
7929      AddNewLine Block, "r", "CONSTRAINT BOUNDS"
          
          ' Apply bounds according to the relation type
          Dim i As Long, BoundType As Long, Comment As String, bound As Double
7930      For i = 1 To n_con
7931          If i Mod 100 = 1 Then
7932              Application.StatusBar = "OpenSolver: Creating .nl file. Writing constraint bounds " & i & "/" & n_con
7933          End If
7934          DoEvents
              
7935          bound = LinearConstants(ConstraintIndexToTreeIndex(i - 1))
7936          ConvertConstraintToNL ConstraintRelations(ConstraintIndexToTreeIndex(i - 1)), BoundType, Comment
7937          AddNewLine Block, BoundType & " " & bound, "    " & ConstraintMapRev(CStr(i - 1)) & Comment & bound
7938      Next i
          
7939      MakeRBlock = StripTrailingNewline(Block)
End Function

' Writes B block for .nl file. This contains the variable bounds
Function MakeBBlock() As String
          Dim Block As String
7940      Block = ""
          
          ' Write block header
7941      AddNewLine Block, "b", "VARIABLE BOUNDS"
          
          Dim i As Long, bound As String, Comment As String, VariableIndex As Long, VarName As String, value As Double
7942      For i = 1 To n_var
7943          If i Mod 100 = 1 Then
7944              Application.StatusBar = "OpenSolver: Creating .nl file. Writing variable bounds " & i & "/" & n_var
7945          End If
7946          DoEvents
              
7947          VariableIndex = VariableNLIndexToCollectionIndex(i - 1)
7948          Comment = "    " & VariableMapRev(CStr(i - 1))
           
7949          If VariableIndex <= numActualVars Then
                  If BinaryVars(VariableIndex) Then
                    bound = "0 0 1"
                    Comment = Comment & " IN [0, 1]"
                  ' Real variables, use actual bounds
7950              ElseIf m.AssumeNonNegative Then
7951                  VarName = m.GetAdjCellName(CLng(VariableIndex))
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
End Function

' Writes K block for .nl file. This contains the cumulative count of non-zero jacobian entries for the first n-1 variables
Function MakeKBlock() As String
          Dim Block As String
7970      Block = ""

          ' Add block header
7971      AddNewLine Block, "k" & n_var - 1, "NUMBER OF JACOBIAN ENTRIES (CUMULATIVE) FOR FIRST " & n_var - 1 & " VARIABLES"
          
          ' Loop through first n_var - 1 variables and add the non-zero count to the running total
          Dim i As Long, total As Long
7972      total = 0
7973      For i = 1 To n_var - 1
7974          If i Mod 100 = 1 Then
7975              Application.StatusBar = "OpenSolver: Creating .nl file. Writing jacobian counts " & i & "/" & n_var - 1
7976          End If
7977          DoEvents
              
7978          total = total + NonZeroConstraintCount(VariableNLIndexToCollectionIndex(i - 1))
7979          AddNewLine Block, CStr(total), "    Up to " & VariableMapRev(CStr(i - 1)) & ": " & CStr(total) & " entries in Jacobian"
7980      Next i
          
7981      MakeKBlock = StripTrailingNewline(Block)
End Function

' Writes J blocks for .nl file. These contain the linear part of each constraint
Function MakeJBlocks() As String
          Dim Block As String
7982      Block = ""
          
          Dim ConstraintElements() As String
          Dim CommentElements() As String
          
          Dim i As Long, TreeIndex As Long, VariableIndex As Long
7983      For i = 1 To n_con
7984          If i Mod 100 = 1 Then
7985              Application.StatusBar = "OpenSolver: Creating .nl file. Writing linear constraints" & i & "/" & n_con
7986          End If
7987          DoEvents
              
7988          TreeIndex = ConstraintIndexToTreeIndex(i - 1)
          
              ' Make header
7989          AddNewLine Block, "J" & i - 1 & " " & LinearConstraints(TreeIndex).NumPresent, "CONSTRAINT LINEAR SECTION " & ConstraintMapRev(i)
              
7990          ReDim ConstraintElements(n_var)
7991          ReDim CommentElements(n_var)
              
              ' Note that the LinearConstraintNL object store the variables in parsed order, but we need to output in .nl order
              ' First we collect all the constraint elements for the variables that are present
              Dim j As Long
7992          For j = 1 To LinearConstraints(TreeIndex).Count
7993              If LinearConstraints(TreeIndex).VariablePresent(j) Then
7994                  VariableIndex = VariableCollectionIndexToNLIndex(j)
7995                  ConstraintElements(VariableIndex) = VariableIndex & " " & LinearConstraints(TreeIndex).Coefficient(j)
7996                  CommentElements(VariableIndex) = "    + " & LinearConstraints(TreeIndex).Coefficient(j) & " * " & VariableMapRev(CStr(VariableIndex))
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
End Function

' Writes the G blocks for .nl file. These contain the linear parts of each objective
Function MakeGBlocks() As String
          Dim Block As String
8006      Block = ""
          
          Dim i As Long, ObjectiveVariables As Collection, ObjectiveCoefficients As Collection
8007      For i = 1 To n_obj
              ' Make header
8008          AddNewLine Block, "G" & i - 1 & " " & LinearObjectives(i).NumPresent, "OBJECTIVE LINEAR SECTION " & ObjectiveCells(i)
              
              ' This loop is not in the right order (see J blocks)
              ' Since the objective only containts one variable, the output will still be correct without reordering
              Dim j As Long, VariableIndex As Long
8009          For j = 1 To LinearObjectives(i).Count
8010              If LinearObjectives(i).VariablePresent(j) Then
8011                  VariableIndex = VariableCollectionIndexToNLIndex(j)
8012                  AddNewLine Block, VariableIndex & " " & LinearObjectives(i).Coefficient(j), "    + " & LinearObjectives(i).Coefficient(j) & " * " & VariableMapRev(CStr(VariableIndex))
8013              End If
8014          Next j
8015      Next i
          
8016      MakeGBlocks = StripTrailingNewline(Block)
End Function

' Writes the .col summary file. This contains the variable names listed in .nl order
Sub OutputColFile()
          Dim ColFilePathName As String
8017      ColFilePathName = GetTempFilePath("model.col")
          
8018      DeleteFileAndVerify ColFilePathName, "Writing Col File", "Couldn't delete the .col file: " & ColFilePathName
          
8019      On Error GoTo ErrHandler
8020      Open ColFilePathName For Output As #2
          
          Dim var As Variant
8021      For Each var In VariableMap
8022          Application.StatusBar = "OpenSolver: Creating .nl file. Writing col file"
8023          DoEvents
              
8024          WriteToFile 2, VariableMapRev(var)
8025      Next var
          
8026      Close #2
8027      Exit Sub
          
ErrHandler:
8028      Close #2
8029      Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
End Sub

' Writes the .row summary file. This contains the constraint names listed in .nl order
Sub OutputRowFile()
          Dim RowFilePathName As String
8030      RowFilePathName = GetTempFilePath("model.row")
          
8031      DeleteFileAndVerify RowFilePathName, "Writing Row File", "Couldn't delete the .row file: " & RowFilePathName
          
8032      On Error GoTo ErrHandler
8033      Open RowFilePathName For Output As #3
          
          Dim con As Variant
8034      For Each con In ConstraintMap
8035          Application.StatusBar = "OpenSolver: Creating .nl file. Writing con file"
8036          DoEvents
              
8037          WriteToFile 3, ConstraintMapRev(CStr(con))
8038      Next con
          
8039      Close #3
8040      Exit Sub
          
ErrHandler:
8041      Close #3
8042      Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
End Sub

Public Sub OutputOptionsFile(OptionsFilePath As String, SolveOptions As SolveOptionsType)
    On Error GoTo ErrHandler
    
    DeleteFileAndVerify OptionsFilePath, "Writing Options File", "Couldn't delete the .opt file: " & OptionsFilePath
    
    Open OptionsFilePath For Output As 4
    Print #4, "iteration_limit " & SolveOptions.MaxIterations
    Print #4, "allowable_fraction_gap " & SolveOptions.Tolerance
    Print #4, "time_limit " & SolveOptions.maxTime
    Close #4
    Exit Sub
    
ErrHandler:
    Close #4
    Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
End Sub

' Adds a new line to the current string, appending LineText at position 0 and CommentText at position CommentSpacing
Sub AddNewLine(CurText As String, LineText As String, Optional CommentText As String = "")
          Dim Comment As String
8043      Comment = ""
          
          ' Add comment with padding if coment should be included
8044      If WriteComments And CommentText <> "" Then
              Dim j As Long
8045          For j = 1 To CommentSpacing - Len(LineText)
8046             Comment = Comment + " "
8047          Next j
8048          Comment = Comment + "# " + CommentText
8049      End If
          
8050      CurText = CurText & LineText & Comment & vbNewLine
End Sub

' Removes a "\n" character from the end of a string
Function StripTrailingNewline(Block As String) As String
8051      If Len(Block) > 1 Then
8052          StripTrailingNewline = left(Block, Len(Block) - 2)
8053      Else
8054          StripTrailingNewline = Block
8055      End If
End Function

Function ConvertFormulaToExpressionTree(strFormula As String) As ExpressionTree
      ' Converts a string formula to a complete expression tree
      ' Uses the Shunting Yard algorithm (adapted to produce an expression tree) which takes O(n) time
      ' https://en.wikipedia.org/wiki/Shunting-yard_algorithm
      ' For details on modifications to algorithm
      ' http://wcipeg.com/wiki/Shunting_yard_algorithm#Conversion_into_syntax_tree
          Dim tksFormula As Tokens
8056      Set tksFormula = ParseFormula("=" + strFormula)
          
          Dim Operands As New ExpressionTreeStack, Operators As New StringStack, ArgCounts As New OperatorArgCountStack
          
          Dim i As Long, c As Range, tkn As Token, tknOld As String, Tree As ExpressionTree
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
8075                      GoTo Mismatch
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
8100                      GoTo Mismatch
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
8113              GoTo Mismatch
8114          End If
              ' Pop the operator onto the operand stack as a new tree
8115          PopOperator Operators, Operands
8116      Loop
          
          ' We are left with a single tree in the operand stack - this is the complete expression tree
8117      Set ConvertFormulaToExpressionTree = Operands.Pop
8118      Exit Function
          
Mismatch:
8119      MsgBox "Mismatched parentheses"
          
End Function

' Creates an ExpressionTree and initialises NodeText and NodeType
Public Function CreateTree(NodeText As String, NodeType As Long) As ExpressionTree
          Dim obj As ExpressionTree

8120      Set obj = New ExpressionTree
8121      obj.NodeText = NodeText
8122      obj.NodeType = NodeType

8123      Set CreateTree = obj
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
Function NumberOfOperands(FunctionName As String, Optional ArgCount As Long = 0) As Long
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
End Function

' Converts common Excel functinos to .nl operators
Function ConvertExcelFunctionToNL(FunctionName As String) As String
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
End Function

' Determines the precedence of arithmetic operators
Function Precedence(tkn As String) As Long
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
Function CheckPrecedence(tkn1 As String, tkn2 As String) As Boolean
          ' Either tkn1 is left-associative and its precedence is less than or equal to that of tkn2
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
End Function

' Determines the left-associativity of arithmetic operators
Function OperatorIsLeftAssociative(tkn As String) As Boolean
8201      Select Case tkn
          Case "plus", "minus", "mult", "div", "eq", "ne", "gt", "ge", "lt", "le"
8202          OperatorIsLeftAssociative = True
8203      Case "pow"
8204          OperatorIsLeftAssociative = False
8205      Case Else
8206          MsgBox "unknown associativity: " & tkn
8207      End Select
End Function

' Pops an operator from the operator stack along with the corresponding number of operands.
Sub PopOperator(Operators As StringStack, Operands As ExpressionTreeStack, Optional ArgCount As Long = 0)
          ' Pop the operator and create a new ExpressionTree
          Dim Operator As String
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
End Sub

' Check whether a token on the operator stack is a function operator (vs. an arithmetic operator)
Function IsFunctionOperator(tkn As String) As Boolean
8216      Select Case tkn
          Case "plus", "minus", "mult", "div", "pow", "neg", "("
8217          IsFunctionOperator = False
8218      Case Else
8219          IsFunctionOperator = True
8220      End Select
End Function

' Negates a tree by adding a 'neg' node to the root
Sub AddNegToTree(Tree As ExpressionTree)
          Dim NewTree As ExpressionTree
8221      Set NewTree = CreateTree("neg", ExpressionTreeOperator)
          
8222      Tree.NodeText = right(Tree.NodeText, Len(Tree.NodeText) - 1)
8223      NewTree.SetChild 1, Tree
          
8224      Set Tree = NewTree
End Sub

' Formats an expression tree node's text as .nl output
Function FormatNL(NodeText As String, NodeType As ExpressionTreeNodeType) As String
8225      Select Case NodeType
          Case ExpressionTreeVariable
8226          FormatNL = "v" & VariableMap(NodeText)
8227      Case ExpressionTreeNumber
8228          FormatNL = "n" & NodeText
8229      Case ExpressionTreeOperator
8230          FormatNL = "o" & CStr(ConvertOperatorToNLCode(NodeText))
8231      End Select
End Function

' Converts an operator string to .nl code
Function ConvertOperatorToNLCode(FunctionName As String) As Long
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
Function ConvertObjectiveSenseToNL(ObjectiveSense As ObjectiveSenseType) As Long
8341      Select Case ObjectiveSense
          Case ObjectiveSenseType.MaximiseObjective
8342          ConvertObjectiveSenseToNL = 1
8343      Case ObjectiveSenseType.MinimiseObjective
8344          ConvertObjectiveSenseToNL = 0
8345      Case Else
8346          MsgBox "Objective sense not supported: " & ObjectiveSense
8347      End Select
End Function

' Converts RelationConsts enum to .nl code.
Sub ConvertConstraintToNL(Relation As RelationConsts, BoundType As Long, Comment As String)
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

