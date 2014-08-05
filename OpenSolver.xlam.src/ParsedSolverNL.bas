Attribute VB_Name = "ParsedSolverNL"
Option Explicit

' This module is for writing .nl files that describe the model and solving these
' For more info on .nl file format see:
' http://citeseerx.ist.psu.edu/viewdoc/summary?doi=10.1.1.60.9659
' http://www.ampl.com/REFS/hooking2.pdf

Dim m As CModelParsed

Dim problem_name As String

Public WriteComments As Boolean     ' Whether .nl file should include comments
Public CommentIndent As Integer     ' Tracks the level of indenting in comments on nl output
Public Const CommentSpacing = 24    ' The column number at which nl comments begin

' ==========================================================================
' ASL variables
' These are variables used in the .NL file that are also used by the AMPL Solver Library
' We use the same names as the ASL for consistency
' ASL header definitions available at http://www.netlib.org/ampl/solvers/asl.h

Dim n_var As Integer            ' Number of variables
Dim n_con As Integer            ' Number of constraints
Dim n_obj As Integer            ' Number of objectives
Dim nranges As Integer          ' Number of range constraints
Dim n_eqn_ As Integer           ' Number of equality constraints
Dim n_lcon As Integer           ' Number of logical constraints

Dim nlc As Integer              ' Number of non-linear constraints
Dim nlo As Integer              ' Number of non-linear objectives

Dim nlnc As Integer             ' Number of non-linear network constraints
Dim lnc As Integer              ' Number of linear network constraints

Dim nlvc As Integer             ' Number of variables appearing non-linearly in constraints
Dim nlvo As Integer             ' Number of variables appearing non-linearly in objectives
Dim nlvb As Integer             ' Number of variables appearing non-linearly in both constraints and objectives

Dim nwv_ As Integer             ' Number of linear network variables
Dim nfunc_ As Integer           ' Number of user-defined functions
Dim arith As Integer            ' Not sure what this does, keep at 0. This flag may indicate whether little- or big-endian arithmetic was used to write the file.
Dim flags As Integer            ' 1 = want output suffixes

Dim nbv As Integer              ' Number of binary variables
Dim niv As Integer              ' Number of other integer variables
Dim nlvbi As Integer            ' Number of integer variables appearing non-linearly in both constraints and objectives
Dim nlvci As Integer            ' Number of integer variables appearing non-linearly just in constraints
Dim nlvoi As Integer            ' Number of integer variables appearing non-linearly just in objectives

Dim nzc As Integer              ' Number of non-zeros in the Jacobian matrix
Dim nzo As Integer              ' Number of non-zeros in objective gradients

Dim maxrownamelen_ As Integer   ' Length of longest constraint name
Dim maxcolnamelen_ As Integer   ' Length of longest variable name

' Common expressions (from "defined" vars) - whatever that means!? We just set them to zero as we don't use defined vars
Dim comb As Integer
Dim comc As Integer
Dim como As Integer
Dim comc1 As Integer
Dim como1 As Integer

' End ASL variables
' ===========================================================================

' The following variables are used for storing the NL model
' Definitions for some of the terms used:
'   - parsed variable index/order:      the order that variables are read into the model. There are numActualVars adjustable cells, followed by numFakeVars formulae variables
'   - .nl variable index/order:         the order that variables are arranged in the .nl output. This follows a strict specification (see Sub MakeVariableMap)
'   - parsed constraint index/order:    the order that constraints are read into the model. There are numActualCons real constraints followed by numFakeCons formulae constraints
'   - .nl constraint index/order:       the order that constraints are arranged in the .nl output. This follows a strict specification (see Sub MakeConstraintMap)
'   - parsed objective index/order:     there is currently only a single objective supported, so this is irrelevant

Dim VariableMap As Collection       ' A map from variable name (e.g. Test1_D4) to .nl variable index (0 to n_var - 1)
Dim VariableMapRev As Collection    ' A map from .nl variable index (0 to n_var - 1) to variable name (e.g. Test1_D4)
Public VariableIndex As Collection  ' A map from variable name (e.g. Test1_D4) to parsed variable index (1 to n_var)

Dim ConstraintMap As Collection     ' A map from constraint name (e.g. c1_Test1_D4) to .nl constraint index (0 to n_con - 1)
Dim ConstraintMapRev As Collection  ' A map from .nl constraint index (0 to n_con - 1) to constraint name (e.g. c1_Test1_D4)

Dim NonLinearConstraintTrees As Collection  ' Collection containing all non-linear constraint ExpressionTrees stored in parsed constraint order
Dim NonLinearObjectiveTrees As Collection   ' Collection containing all non-linear objective ExpressionTrees stored in parsed objective order

Dim NonLinearVars() As Boolean          ' Array of size n_var indicating whether each variable appears non-linearly in the model
Dim NonLinearConstraints() As Boolean   ' Array of size n_con indicating whether each constraint has non-linear elements

Dim ConstraintIndexToTreeIndex() As Integer         ' Array of size n_con storing the parsed constraint index for each .nl constraint index
Dim VariableNLIndexToCollectionIndex() As Integer   ' Array of size n_var storing the parsed variable index for each .nl variable index
Dim VariableCollectionIndexToNLIndex() As Integer   ' Array of size n_var storing the .nl variable index for each parsed variable index

Dim LinearConstraints As Collection     ' Collection containing the LinearConstraintNLs for each constraint stored in parsed constraint order
Dim LinearConstants As Collection       ' Collection containing the constant (a Double) for each constraint stored in parsed constraint order
Dim LinearObjectives As Collection      ' Collection containing the LinearConstraintNLs for each objective stored in parsed objective order

Dim ConstraintRelations As Collection   ' Collection containing the RelationConst for each constraint stored in parsed constraint order

Dim ObjectiveCells As Collection    ' Collection containing the objective cells for each objective stored in parsed objective order
Dim ObjectiveSenses As Collection   ' Collection containing the objective sense for each objective stored in parsed objective order

Dim numActualVars As Integer    ' The number of actual variables in the parsed model (the adjustable cells in the Solver model)
Dim numFakeVars As Integer      ' The number of "fake" variables in the parsed model (variables that arise from parsing the formulae in the Solver model)
Dim numActualCons As Integer    ' The number of actual constraints in the parsed model (the constraints defined in the Solver model)
Dim numFakeCons As Integer      ' The number of "fake" constraints in the parsed model (the constraints that arise from parsing the formulae in the Solver model)
Dim numActualEqs As Integer     ' The number of actual equality constraints in the parsed model
Dim numActualRanges As Integer  ' The number of actual inequality constraints in the parsed model

Dim NonZeroConstraintCount() As Integer ' An array of size n_var counting the number of times that each variable (in parsed variable order) appears in the linear constraints

' Public accessor for VariableCollectionIndexToNLIndex
Public Property Get GetVariableNLIndex(index As Integer) As Integer
    GetVariableNLIndex = VariableCollectionIndexToNLIndex(index)
End Property

' Creates .nl file and solves model
Function SolveModelParsed_NL(ModelFilePathName As String, model As CModelParsed, Optional ShouldWriteComments As Boolean = True)
    Set m = model
    WriteComments = ShouldWriteComments
    
    ' =============================================================
    ' Process model for .nl output
    ' All Module-level variables required for .nl output should be set in this step
    ' No modification to these variables should be done while writing the .nl file
    ' =============================================================
    
    InitialiseModelStats
    
    CreateVariableIndex
    
    ProcessFormulae
    ProcessObjective
    
    MakeVariableMap
    MakeConstraintMap
    
    ' =============================================================
    ' Write output files
    ' =============================================================
    
    ' Create supplementary outputs
    OutputColFile
    OutputRowFile
    
    ' Write .nl file
    On Error GoTo ErrHandler
    
    Open ModelFilePathName For Output As #1
    
    ' Write header
    Print #1, MakeHeader()
    ' Write C blocks
    Print #1, MakeCBlocks()
    ' Write O block
    Print #1, MakeOBlocks()
    ' Write d block
    Print #1, MakeDBlock()
    ' Write x block
    Print #1, MakeXBlock()
    ' Write r block
    Print #1, MakeRBlock()
    ' Write b block
    Print #1, MakeBBlock()
    ' Write k block
    Print #1, MakeKBlock()
    ' Write J block
    Print #1, MakeJBlocks()
    ' Write G block
    Print #1, MakeGBlocks()
    
    Close #1
    
    ' =============================================================
    ' Solve model using chosen solver
    ' =============================================================
    
    Dim SolutionFilePathName As String
    SolutionFilePathName = SolutionFilePath(m.Solver)
    
    Dim ExternalSolverPathName As String
    ExternalSolverPathName = CreateSolveScriptParsed(m.Solver, ModelFilePathName)
             
    Dim logCommand As String, logFileName As String
'    logFileName = "log1.tmp"
'    logCommand = " > " & """" & ConvertHfsPath(GetTempFolder) & logFileName & """"
                  
    Dim ExecutionCompleted As Boolean
    ExternalSolverPathName = """" & ConvertHfsPath(ExternalSolverPathName) & """"
              
    Dim exeResult As Long, userCancelled As Boolean
    ExecutionCompleted = OSSolveSync(ExternalSolverPathName, "", "", logCommand, SW_SHOWNORMAL, True, userCancelled, exeResult) ' Run solver, waiting for completion
    If userCancelled Then
        ' User pressed escape. Dialogs have already been shown. Exit with a 'cancelled' error
        On Error GoTo ErrHandler
        Err.Raise Number:=OpenSolver_UserCancelledError, Source:="Solving NL model", Description:="The solving process was cancelled by the user."
    End If
    If exeResult <> 0 Then
        ' User pressed escape. Dialogs have already been shown. Exit with a 'cancelled' error
        On Error GoTo ErrHandler
        Err.Raise Number:=OpenSolver_SolveError, Source:="Solving NL model", Description:="The " & m.Solver & " solver did not complete, but aborted with the error code " & exeResult & "." & vbCrLf & vbCrLf & "The last log file can be viewed under the OpenSolver menu and may give you more information on what caused this error."
    End If
    
    ' =============================================================
    ' Read results from solution file
    ' =============================================================
    
    If Not FileOrDirExists(SolutionFilePathName) Then
        On Error GoTo ErrHandler
        Err.Raise Number:=OpenSolver_SolveError, Source:="Solving NL model", Description:="The solver did not create a solution file. No new solution is available."
    End If
    Dim solutionLoaded As Boolean, errorString As String
    solutionLoaded = ReadModelParsed(m.Solver, SolutionFilePathName, errorString, m)
    On Error GoTo ErrHandler
    If errorString <> "" Then
        Err.Raise Number:=OpenSolver_SolveError, Source:="Solving NL model", Description:=errorString
    ElseIf Not solutionLoaded Then 'read error
        SolveModelParsed_NL = False
    End If

    SolveModelParsed_NL = True
    Exit Function
    
ErrHandler:
    Close #1
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Sub InitialiseModelStats()
    ' Number of actual variables is the number of adjustable cells in the Solver model
    numActualVars = m.AdjustableCells.Count
    ' Number of fake variables is the number of formulae equations we have created
    numFakeVars = m.Formulae.Count
    ' Number of actual constraints is the number of constraints in the Solver model
    numActualCons = m.LHSKeys.Count
    ' Number of fake constraints is the number of formulae equations we have created
    numFakeCons = m.Formulae.Count
    
    ' Divide the actual constraints into equalities and inequalities (ranges)
    Dim i As Integer
    For i = 1 To numActualCons
        If m.Rels(i) = RelationConsts.RelationEQ Then
            numActualEqs = numActualEqs + 1
        Else
            numActualRanges = numActualRanges + 1
        End If
    Next i
    
    ' ===============================================================================
    ' Initialise the ASL variables - see definitions for explanation of each variable
    
    ' Model statistics for line #1
    problem_name = "'Sheet=" + m.SolverModelSheet.Name + "'"
    
    ' Model statistics for line #2
    n_var = numActualVars + numFakeVars
    n_con = numActualCons + numFakeCons
    n_obj = 0
    nranges = numActualRanges
    n_eqn_ = numActualEqs + numFakeCons     ' All fake formulae constraints are equalities
    n_lcon = 0
    
    ' Model statistics for line #3
    nlc = 0
    nlo = 0
    
    ' Model statistics for line #4
    nlnc = 0
    lnc = 0
    
    ' Model statistics for line #5
    nlvc = 0
    nlvo = 0
    nlvb = 0
    
    ' Model statistics for line #6
    nwv_ = 0
    nfunc_ = 0
    arith = 0
    flags = 0
    
    ' Model statistics for line #7
    nbv = 0
    niv = 0
    nlvbi = 0
    nlvci = 0
    nlvoi = 0
    
    ' Model statistics for line #8
    nzc = 0
    nzo = 0
    
    ' Model statistics for line #9
    maxrownamelen_ = 0
    maxcolnamelen_ = 0
    
    ' Model statistics for line #10
    comb = 0
    comc = 0
    como = 0
    comc1 = 0
    como1 = 0
End Sub

' Creates map from variable name (e.g. Test1_D4) to parsed variable index (1 to n_var)
Sub CreateVariableIndex()
    Set VariableIndex = New Collection
    Dim c As Range, cellName As String, i As Integer
    
    
    ' First read in actual vars
    i = 1
    For Each c In m.AdjustableCells
        cellName = ConvertCellToStandardName(c)
        
        ' Update variable maps
        VariableIndex.Add i, cellName
        i = i + 1
    Next c
    
    ' Next read in fake formulae vars
    For i = 1 To numFakeVars
        cellName = m.Formulae(i).strAddress
        
        ' Update variable maps
        VariableIndex.Add i + numActualVars, cellName
    Next i
End Sub

' Creates maps from variable name (e.g. Test1_D4) to .nl variable index (0 to n_var - 1) and vice-versa, and
' maps from parsed variable index to .nl variable index and vice-versa
Sub MakeVariableMap()
    ' =============================================
    ' Create index of variable names in parsed variable order
    Dim CellNames As New Collection
    
    ' Actual variables
    Dim c As Range, cellName As String, i As Integer
    For Each c In m.AdjustableCells
        cellName = ConvertCellToStandardName(c)
        CellNames.Add cellName
    Next c
    
    ' Formulae variables
    For i = 1 To m.Formulae.Count
        cellName = m.Formulae(i).strAddress
        CellNames.Add cellName
    Next i
    
    ' ==============================================
    ' Add variables to the variable map in the required order
    ReDim VariableNLIndexToCollectionIndex(n_var)
    ReDim VariableCollectionIndexToNLIndex(n_var)
    Set VariableMap = New Collection
    Set VariableMapRev = New Collection
    
    Dim index As Integer
    index = 0
    
    ' We loop through the variables and arrange them in the required order:
    '     1st - non-linear continuous
    '     2nd - non-linear integer
    '     3rd - linear arcs (N/A)
    '     4th - other linear
    '     5th - binary
    '     6th - other integer
    
    ' Non-linear
    '    continuous
    '    integer
    For i = 1 To n_var
        If NonLinearVars(i) Then
            'TODO split integer
            AddVariable CellNames(i), index, i
        End If
    Next i
    
    ' Linear
    '    continuous
    '    binary
    '    integer
    For i = 1 To n_var
        If Not NonLinearVars(i) Then
            'TODO split integer/binary
            AddVariable CellNames(i), index, i
        End If
    Next i
    
End Sub

' Adds a variable to all the variable maps with:
'   variable name:            cellName
'   .nl variable index:       index
'   parsed variable index:    i
Sub AddVariable(cellName As String, index As Integer, i As Integer)
    ' Update variable maps
    VariableMap.Add CStr(index), cellName
    VariableMapRev.Add cellName, CStr(index)
    VariableNLIndexToCollectionIndex(index) = i
    VariableCollectionIndexToNLIndex(i) = index
    
    ' Update max length of variable name
    If Len(cellName) > maxcolnamelen_ Then
       maxcolnamelen_ = Len(cellName)
    End If
    
    ' Increase index for the next variable
    index = index + 1
End Sub

' Creates maps from constraint name (e.g. c1_Test1_D4) to .nl constraint index (0 to n_con - 1) and vice-versa, and
' map from .nl constraint index to parsed constraint index
Sub MakeConstraintMap()
    
    Set ConstraintMap = New Collection
    Set ConstraintMapRev = New Collection
    ReDim ConstraintIndexToTreeIndex(n_con)
    
    Dim index As Integer, i As Integer, cellName As String
    index = 0
    
    ' We loop through the constraints and arrange them in the required order:
    '     1st - non-linear
    '     2nd - non-linear network (N/A)
    '     3rd - linear network (N/A)
    '     4th - linear
    
    ' Non-linear constraints
    For i = 1 To n_con
        If NonLinearConstraints(i) Then
            ' Actual constraints
            If i <= numActualCons Then
                cellName = "c" & i & "_" & m.LHSKeys(i)
            ' Formulae constraints
            Else
                cellName = "f" & i & "_" & m.Formulae(i - numActualCons).strAddress
            End If
            AddConstraint cellName, index, i
        End If
    Next i
    
    ' Linear constraints
    For i = 1 To n_con
        If Not NonLinearConstraints(i) Then
            ' Actual constraints
            If i <= numActualCons Then
                cellName = "c" & i & "_" & m.LHSKeys(i)
            ' Formulae constraints
            Else
                cellName = "f" & i & "_" & m.Formulae(i - numActualCons).strAddress
            End If
            AddConstraint cellName, index, i
        End If
    Next i
End Sub

' Adds a constraint to all the constraint maps with:
'   constraint name:          cellName
'   .nl constraint index:     index
'   parsed constraint index:  i
Sub AddConstraint(cellName As String, index As Integer, i As Integer)
    ' Update constraint maps
    ConstraintMap.Add index, cellName
    ConstraintMapRev.Add cellName, CStr(index)
    ConstraintIndexToTreeIndex(index) = i
    
    ' Update max length
    If Len(cellName) > maxrownamelen_ Then
       maxrownamelen_ = Len(cellName)
    End If
    
    ' Increase index for next constraint
    index = index + 1
End Sub

' Processes all constraint formulae into the .nl model formats
Sub ProcessFormulae()
    
    Set NonLinearConstraintTrees = New Collection
    Set LinearConstraints = New Collection
    Set LinearConstants = New Collection
    Set ConstraintRelations = New Collection
    
    ReDim NonLinearVars(n_var)
    ReDim NonLinearConstraints(n_con)
    ReDim NonZeroConstraintCount(n_var)
    
    ' Loop through all constraints and process each
    Dim i As Integer
    For i = 1 To numActualCons
        ProcessSingleFormula m.RHSKeys(i), m.LHSKeys(i), m.Rels(i)
    Next i
    For i = 1 To numFakeCons
        ProcessSingleFormula m.Formulae(i).strFormulaParsed, m.Formulae(i).strAddress, RelationConsts.RelationEQ
    Next i
End Sub

' Processes a single constraint into .nl format. We require:
'     - a non-linear ExpressionTree for all non-linear parts of the equation
'     - a linear LinearConstraintNL for the linear parts of the equation
'     - a constant Double for the constant part of the equation
' We also use the results of processing to update some of the model statistics
Sub ProcessSingleFormula(RHSExpression As String, LHSVariable As String, Relation As RelationConsts)
    ' Convert the string formula into an ExpressionTree object
    Dim Tree As ExpressionTree
    Set Tree = ConvertFormulaToExpressionTree(RHSExpression)
    
    ' The .nl file needs a linear coefficient for every variable in the constraint - non-linear or otherwise
    ' We need a list of all variables in this constraint so that we can know to include them in the linear part of the constraint.
    Dim constraint As New LinearConstraintNL
    constraint.Count = n_var
    Tree.ExtractVariables constraint

    ' Remove linear terms from non-linear trees
    Dim LinearTrees As New Collection
    Tree.MarkLinearity
    Tree.PruneLinearTrees LinearTrees, True
    
    ' Process linear trees to separate constants and variables
    Dim constant As Double
    constant = 0
    Dim j As Integer
    For j = 1 To LinearTrees.Count
        LinearTrees(j).ConvertLinearTreeToConstraint constraint, constant
    Next j
    
    ' Our constraints all have a single term on the LHS and a formulae on the right.
    ' The single LHS term needs to be included in the expression as a linear term with coefficient 1
    ' Move variable from the LHS to the linear constraint with coefficient -1
    constraint.VariablePresent(VariableIndex(LHSVariable)) = True
    constraint.Coefficient(VariableIndex(LHSVariable)) = constraint.Coefficient(VariableIndex(LHSVariable)) - 1
    
    ' Constant must be positive in the final form of the constraint in the .nl file
    If constant < 0 Then
        ' Take constant to the LHS to make it positive
        constant = -constant
        
        ' Our RHS and LHS have now been swapped compared to the original constraint. We need to swap the LE or GE relation if present
        Select Case Relation
        Case RelationConsts.RelationGE
            Relation = RelationLE
        Case RelationConsts.RelationLE
            Relation = RelationGE
        End Select
    Else
        ' Keep constant on RHS and take everything else to LHS.
        ' Need to flip all sign of elements in this constraint
        
        ' Flip all coefficients in linear constraint
        constraint.InvertCoefficients
        
        ' Negate non-linear tree
        Set Tree = Tree.Negate
    End If
    
    
    ' Save results of processing
    NonLinearConstraintTrees.Add Tree
    LinearConstraints.Add constraint
    LinearConstants.Add constant
    ConstraintRelations.Add Relation
    
    ' Mark any non-linear variables that we haven't seen before by extracting any remaining variables from the non-linear tree.
    ' Any variable still present in the constraint must be part of the non-linear section
    Dim TempConstraint As New LinearConstraintNL
    TempConstraint.Count = n_var
    Tree.ExtractVariables TempConstraint
    For j = 1 To TempConstraint.Count
        If TempConstraint.VariablePresent(j) And Not NonLinearVars(j) Then
            NonLinearVars(j) = True
            nlvc = nlvc + 1
        End If
    Next j
    
    ' Increase count of non-linear constraints if the non-linear tree is non-empty
    ' An empty tree has a single "0" node
    If Tree.NodeText <> "0" Then
        Dim ConstraintIndex As Integer
        ConstraintIndex = NonLinearConstraintTrees.Count
        NonLinearConstraints(ConstraintIndex) = True
        nlc = nlc + 1
    End If
    
    ' Update jacobian counts using the linear variables present
    For j = 1 To constraint.Count
        ' The nl documentation says that the jacobian counts relate to "the numbers of nonzeros in the first n var - 1 columns of the Jacobian matrix"
        ' This means we should increases the count for any variable that is present and has a non-zero coefficient
        ' However, .nl files generated by AMPL seem to increase the count for any present variable, even if the coefficient is zero (i.e. non-linear variables)
        ' We adopt the AMPL behaviour as it gives faster solve times (see Test 40). Swap the If conditions below to change this behaviour
        
        'If constraint.VariablePresent(j) And constraint.Coefficient(j) <> 0 Then
        If constraint.VariablePresent(j) Then
            NonZeroConstraintCount(j) = NonZeroConstraintCount(j) + 1
            nzc = nzc + 1
        End If
    Next j
End Sub

' Process objective function into .nl format. We require the same as for constraints:
'     - a non-linear ExpressionTree for all non-linear parts of the equation
'     - a linear LinearConstraintNL for the linear parts of the equation
'     - a constant Double for the constant part of the equation
Sub ProcessObjective()
    Set NonLinearObjectiveTrees = New Collection
    Set ObjectiveSenses = New Collection
    Set ObjectiveCells = New Collection
    Set LinearObjectives = New Collection
    
    
    ' Currently just one objective - a single linear variable
    ObjectiveCells.Add ConvertCellToStandardName(m.ObjectiveCell)
    ObjectiveSenses.Add m.ObjectiveSense
    
    ' Objective non-linear constraint tree is empty
    NonLinearObjectiveTrees.Add CreateTree("0", ExpressionTreeNodeType.ExpressionTreeNumber)
    
    ' Objective has a single linear term - the objective variable with coefficient 1
    Dim Objective As New LinearConstraintNL
    Objective.Count = n_var
    Objective.VariablePresent(VariableIndex(ObjectiveCells(1))) = True
    Objective.Coefficient(VariableIndex(ObjectiveCells(1))) = 1
    
    ' Save results
    LinearObjectives.Add Objective
    
    ' Track non-zero jacobian count in objective
    nzo = nzo + 1
    
    ' Adjust objective count
    n_obj = n_obj + 1
End Sub

' Writes header block for .nl file. This contains the model statistics
Function MakeHeader() As String
    Dim Header As String
    Header = ""
    
    'Line #1
    AddNewLine Header, "g3 1 1 0", "problem " & problem_name
    'Line #2
    AddNewLine Header, " " & n_var & " " & n_con & " " & n_obj & " " & nranges & " " & n_eqn_ & " " & n_lcon, "vars, constraints, objectives, ranges, eqns"
    'Line #3
    AddNewLine Header, " " & nlc & " " & nlo, "nonlinear constraints, objectives"
    'Line #4
    AddNewLine Header, " " & nlnc & " " & lnc, "network constraints: nonlinear, linear"
    'Line #5
    AddNewLine Header, " " & nlvc & " " & nlvo & " " & nlvb, "nonlinear vars in constraints, objectives, both"
    'Line #6
    AddNewLine Header, " " & nwv_ & " " & nfunc_ & " " & arith & " " & flags, "linear network variables; functions; arith, flags"
    'Line #7
    AddNewLine Header, " " & nbv & " " & niv & " " & nlvbi & " " & nlvci & " " & nlvoi, "discrete variables: binary, integer, nonlinear (b,c,o)"
    'Line #8
    AddNewLine Header, " " & nzc & " " & nzo, "nonzeros in Jacobian, gradients"
    'Line #9
    AddNewLine Header, " " & maxrownamelen_ & " " & maxcolnamelen_, "max name lengths: constraints, variables"
    'Line #10
    AddNewLine Header, " " & comb & " " & comc & " " & como & " " & comc1 & " " & como1, "common exprs: b,c,o,c1,o1"
    
    MakeHeader = StripTrailingNewline(Header)
End Function

' Writes C blocks for .nl file. These describe the non-linear parts of each constraint.
Function MakeCBlocks() As String
    Dim Block As String
    Block = ""
    
    Dim i As Integer
    For i = 1 To n_con
        ' Add block header for the constraint
        AddNewLine Block, "C" & i - 1, "CONSTRAINT NON-LINEAR SECTION " + ConstraintMapRev(CStr(i - 1))
        
        ' Add expression tree
        CommentIndent = 4
        Block = Block + NonLinearConstraintTrees(ConstraintIndexToTreeIndex(i - 1)).ConvertToNL
    Next i
    
    MakeCBlocks = StripTrailingNewline(Block)
End Function

' Writes O blocks for .nl file. These describe the non-linear parts of each objective.
Function MakeOBlocks() As String
    Dim Block As String
    Block = ""
    
    Dim i As Integer
    For i = 1 To n_obj
        ' Add block header for the objective
        AddNewLine Block, "O" & i - 1 & " " & ConvertObjectiveSenseToNL(ObjectiveSenses(i)), "OBJECTIVE NON-LINEAR SECTION " & ObjectiveCells(i)
        
        ' Add expression tree
        CommentIndent = 4
        Block = Block + NonLinearObjectiveTrees(i).ConvertToNL
    Next i
    
    MakeOBlocks = StripTrailingNewline(Block)
End Function

' Writes D block for .nl file. This contains the initial guess for dual variables.
' We don't use this, so just set them all to zero
Function MakeDBlock() As String
    Dim Block As String
    Block = ""
    
    ' Add block header
    AddNewLine Block, "d" & n_con, "INITIAL DUAL GUESS"
    
    ' Set duals to zero for all constraints
    Dim i As Integer
    For i = 1 To n_con
        AddNewLine Block, i - 1 & " 0", "    " & ConstraintMapRev(CStr(i - 1)) & " = " & 0
    Next i
    
    MakeDBlock = StripTrailingNewline(Block)
End Function

' Writes X block for .nl file. This contains the initial guess for primal variables
Function MakeXBlock() As String
    Dim Block As String
    Block = ""
    
    ' Add block header
    AddNewLine Block, "x" & n_var, "INITIAL PRIMAL GUESS"

    ' Loop through the variables in .nl variable order
    Dim i As Integer, initial As String, VariableIndex As Integer
    For i = 1 To n_var
        VariableIndex = VariableNLIndexToCollectionIndex(i - 1)
        
        ' Get initial values
        If VariableIndex <= numActualVars Then
            ' Actual variables - use the value in the actual cell
            initial = CDbl(m.AdjustableCells(VariableIndex))
        Else
            ' Formulae variables - use the initial value saved in the CFormula instance
            initial = CDbl(m.Formulae(VariableIndex - numActualVars).initialValue)
        End If
        AddNewLine Block, i - 1 & " " & initial, "    " & VariableMapRev(CStr(i - 1)) & " = " & initial
    Next i
    
    MakeXBlock = StripTrailingNewline(Block)
End Function

' Writes R block for .nl file. This contains the constant values for each constraint and the relation type
Function MakeRBlock() As String
    Dim Block As String
    Block = ""
     ' Add block header
    AddNewLine Block, "r", "CONSTRAINT BOUNDS"
    
    ' Apply bounds according to the relation type
    Dim i As Integer, BoundType As Integer, Comment As String, bound As Double
    For i = 1 To n_con
        bound = LinearConstants(ConstraintIndexToTreeIndex(i - 1))
        ConvertConstraintToNL ConstraintRelations(ConstraintIndexToTreeIndex(i - 1)), BoundType, Comment
        AddNewLine Block, BoundType & " " & bound, "    " & ConstraintMapRev(CStr(i - 1)) & Comment & bound
    Next i
    
    MakeRBlock = StripTrailingNewline(Block)
End Function

' Writes B block for .nl file. This contains the variable bounds
Function MakeBBlock() As String
    Dim Block As String
    Block = ""
    
    ' Write block header
    AddNewLine Block, "b", "VARIABLE BOUNDS"
    
    Dim i As Integer, bound As String, Comment As String, VariableIndex As Integer
    For i = 1 To n_var
        VariableIndex = VariableNLIndexToCollectionIndex(i - 1)
        Comment = "    " & VariableMapRev(CStr(i - 1))
     
        If VariableIndex <= numActualVars Then
            ' Real variables, use actual bounds
            ' TODO apply actual bounding
            If m.AssumeNonNegative Then
                bound = "2 0"
                Comment = Comment & " >= 0"
            Else
                bound = "3"
                Comment = Comment & " FREE"
            End If
        Else
            ' Fake formulae variables - no bounds
            bound = "3"
            Comment = Comment & " FREE"
        End If
        AddNewLine Block, bound, Comment
    Next i
    
    MakeBBlock = StripTrailingNewline(Block)
End Function

' Writes K block for .nl file. This contains the cumulative count of non-zero jacobian entries for the first n-1 variables
Function MakeKBlock() As String
    Dim Block As String
    Block = ""

    ' Add block header
    AddNewLine Block, "k" & n_var - 1, "NUMBER OF JACOBIAN ENTRIES (CUMULATIVE) FOR FIRST " & n_var - 1 & " VARIABLES"
    
    ' Loop through first n_var - 1 variables and add the non-zero count to the running total
    Dim i As Integer, total As Integer
    total = 0
    For i = 1 To n_var - 1
        total = total + NonZeroConstraintCount(VariableNLIndexToCollectionIndex(i - 1))
        AddNewLine Block, CStr(total), "    Up to " & VariableMapRev(CStr(i - 1)) & ": " & CStr(total) & " entries in Jacobian"
    Next i
    
    MakeKBlock = StripTrailingNewline(Block)
End Function

' Writes J blocks for .nl file. These contain the linear part of each constraint
Function MakeJBlocks() As String
    Dim Block As String
    Block = ""
    
    Dim ConstraintElements() As String
    Dim CommentElements() As String
    
    Dim i As Integer, TreeIndex As Integer, VariableIndex As Integer
    For i = 1 To n_con
        TreeIndex = ConstraintIndexToTreeIndex(i - 1)
    
        ' Make header
        AddNewLine Block, "J" & i - 1 & " " & LinearConstraints(TreeIndex).NumPresent, "CONSTRAINT LINEAR SECTION " & ConstraintMapRev(i)
        
        ReDim ConstraintElements(n_var)
        ReDim CommentElements(n_var)
        
        ' Note that the LinearConstraintNL object store the variables in parsed order, but we need to output in .nl order
        ' First we collect all the constraint elements for the variables that are present
        Dim j As Integer
        For j = 1 To LinearConstraints(TreeIndex).Count
            If LinearConstraints(TreeIndex).VariablePresent(j) Then
                VariableIndex = VariableCollectionIndexToNLIndex(j)
                ConstraintElements(VariableIndex) = VariableIndex & " " & LinearConstraints(TreeIndex).Coefficient(j)
                CommentElements(VariableIndex) = "    + " & LinearConstraints(TreeIndex).Coefficient(j) & " * " & VariableMapRev(CStr(VariableIndex))
            End If
        Next j
        
        ' Output the constraint elements to the J Block in .nl variable order
        For j = 0 To n_var - 1
            If ConstraintElements(j) <> "" Then
                AddNewLine Block, ConstraintElements(j), CommentElements(j)
            End If
        Next j
        
    Next i
    
    MakeJBlocks = StripTrailingNewline(Block)
End Function

' Writes the G blocks for .nl file. These contain the linear parts of each objective
Function MakeGBlocks() As String
    Dim Block As String
    Block = ""
    
    Dim i As Integer, ObjectiveVariables As Collection, ObjectiveCoefficients As Collection
    For i = 1 To n_obj
        ' Make header
        AddNewLine Block, "G" & i - 1 & " " & LinearObjectives(i).NumPresent, "OBJECTIVE LINEAR SECTION " & ObjectiveCells(i)
        
        ' This loop is not in the right order (see J blocks)
        ' Since the objective only containts one variable, the output will still be correct without reordering
        Dim j As Integer, VariableIndex As Integer
        For j = 1 To LinearObjectives(i).Count
            If LinearObjectives(i).VariablePresent(j) Then
                VariableIndex = VariableCollectionIndexToNLIndex(j)
                AddNewLine Block, VariableIndex & " " & LinearObjectives(i).Coefficient(j), "    + " & LinearObjectives(i).Coefficient(j) & " * " & VariableMapRev(CStr(VariableIndex))
            End If
        Next j
    Next i
    
    MakeGBlocks = StripTrailingNewline(Block)
End Function

' Writes the .col summary file. This contains the variable names listed in .nl order
Sub OutputColFile()
    Dim ColFilePathName As String
    ColFilePathName = GetTempFilePath("model.col")
    
    DeleteFileAndVerify ColFilePathName, "Writing Col File", "Couldn't delete the .col file: " & ColFilePathName
    
    Open ColFilePathName For Output As #2
    
    Dim var As Variant
    For Each var In VariableMap
        WriteToFile 2, VariableMapRev(var)
    Next var
    
    Close #2
End Sub

' Writes the .row summary file. This contains the constraint names listed in .nl order
Sub OutputRowFile()
    Dim RowFilePathName As String
    RowFilePathName = GetTempFilePath("model.row")
    
    DeleteFileAndVerify RowFilePathName, "Writing Row File", "Couldn't delete the .row file: " & RowFilePathName
    
    Open RowFilePathName For Output As #3
    
    Dim con As Variant
    For Each con In ConstraintMap
        WriteToFile 3, ConstraintMapRev(CStr(con))
    Next con
    
    Close #3
End Sub

' Adds a new line to the current string, appending LineText at position 0 and CommentText at position CommentSpacing
Sub AddNewLine(CurText As String, LineText As String, Optional CommentText As String = "")
    Dim Comment As String
    Comment = ""
    
    ' Add comment with padding if coment should be included
    If WriteComments And CommentText <> "" Then
        Dim j As Integer
        For j = 1 To CommentSpacing - Len(LineText)
           Comment = Comment + " "
        Next j
        Comment = Comment + "# " + CommentText
    End If
    
    CurText = CurText & LineText & Comment & vbNewLine
End Sub

' Removes a "\n" character from the end of a string
Function StripTrailingNewline(Block As String) As String
    StripTrailingNewline = left(Block, Len(Block) - 2)
End Function

Function ConvertFormulaToExpressionTree(strFormula As String) As ExpressionTree
' Converts a string formula to a complete expression tree
' Uses the Shunting Yard algorithm (adapted to produce an expression tree) which takes O(n) time
' https://en.wikipedia.org/wiki/Shunting-yard_algorithm
' For details on modifications to algorithm
' http://wcipeg.com/wiki/Shunting_yard_algorithm#Conversion_into_syntax_tree
    Dim tksFormula As Tokens
    Set tksFormula = modTokeniser.ParseFormula("=" + strFormula)
    
    Dim Operands As New ExpressionTreeStack, Operators As New StringStack
    
    Dim i As Integer, c As Range, tkn As Token, tknOld As String, Tree As ExpressionTree
    For i = 1 To tksFormula.Count
        Set tkn = tksFormula.Item(i)
        
        Select Case tkn.TokenType
        ' If the token is a number or variable, then add it to the operands stack as a new tree.
        Case TokenType.Number
            ' Might be a negative number, if so we need to parse out the neg operator
            Set Tree = CreateTree(tkn.Text, ExpressionTreeNumber)
            If left(Tree.NodeText, 1) = "-" Then
                AddNegToTree Tree
            End If
            Operands.Push Tree
            
        Case TokenType.Reference
            Operands.Push CreateTree(tkn.Text, ExpressionTreeVariable)
            
        ' If the token is a function token, then push it onto the operators stack along with a left parenthesis (tokeniser strips the parenthesis).
        Case TokenType.FunctionOpen
            Operators.Push ConvertExcelFunctionToNL(tkn.Text)
            Operators.Push "("
            
        ' If the token is a function argument separator (e.g., a comma)
        Case TokenType.ParameterSeparator
            ' Until the token at the top of the operator stack is a left parenthesis, pop operators off the stack onto the operands stack as a new tree.
            ' If no left parentheses are encountered, either the separator was misplaced or parentheses were mismatched.
            Do While Operators.Peek() <> "("
                PopOperator Operators, Operands
                
                ' If the operator stack runs out without finding a left parenthesis, then there are mismatched parentheses
                If Operators.Count = 0 Then
                    GoTo Mismatch
                End If
            Loop
        
        ' If the token is an operator
        Case TokenType.ArithmeticOperator
            tkn.Text = ConvertExcelFunctionToNL(tkn.Text)
            ' While there is an operator token at the top of the operator stack
            Do While Operators.Count > 0
                tknOld = Operators.Peek()
                ' If either tkn is left-associative and its precedence is less than or equal to that of tknOld
                ' or tkn has precedence less than that of tknOld
                If CheckPrecedence(tkn.Text, tknOld) Then
                    ' Pop tknOld off the operator stack onto the operand stack as a new tree
                    PopOperator Operators, Operands
                Else
                    Exit Do
                End If
            Loop
            ' Push operator onto the operator stack
            Operators.Push tkn.Text
            
        ' If the token is a left parenthesis, then push it onto the operator stack
        Case TokenType.SubExpressionOpen
            Operators.Push tkn.Text
            
        ' If the token is a right parenthesis
        Case TokenType.SubExpressionClose, TokenType.FunctionClose
            ' Until the token at the top of the operator stack is not a left parenthesis, pop operators off the stack onto the operand stack as a new tree.
            Do While Operators.Peek <> "("
                PopOperator Operators, Operands
                ' If the operator stack runs out without finding a left parenthesis, then there are mismatched parentheses
                If Operators.Count = 0 Then
                    GoTo Mismatch
                End If
            Loop
            ' Pop the left parenthesis from the operator stack, but not onto the operand stack
            Operators.Pop
            ' If the token at the top of the stack is a function token, pop it onto the operand stack as a new tree
            If Operators.Count > 0 Then
                If IsFunctionOperator(Operators.Peek()) Then
                    PopOperator Operators, Operands
                End If
            End If
        End Select
    Next i
    
    ' While there are still tokens in the operator stack
    Do While Operators.Count > 0
        ' If the token on the top of the operator stack is a parenthesis, then there are mismatched parentheses
        If Operators.Peek = "(" Then
            GoTo Mismatch
        End If
        ' Pop the operator onto the operand stack as a new tree
        PopOperator Operators, Operands
    Loop
    
    ' We are left with a single tree in the operand stack - this is the complete expression tree
    Set ConvertFormulaToExpressionTree = Operands.Pop
    Exit Function
    
Mismatch:
    MsgBox "Mismatched parentheses"
    
End Function

' Creates an ExpressionTree and initialises NodeText and NodeType
Public Function CreateTree(NodeText As String, NodeType As Long) As ExpressionTree
    Dim obj As ExpressionTree

    Set obj = New ExpressionTree
    obj.NodeText = NodeText
    obj.NodeType = NodeType

    Set CreateTree = obj
End Function

' Determines the number of operands expected by a .nl operator
Function NumberOfOperands(FunctionName As String) As Integer
    Select Case FunctionName
    Case "floor", "ceil", "abs", "neg", "not", "tanh", "tan", "sqrt", "sinh", "sin", "log10", "log", "exp", "cosh", "cos", "atanh", "atan", "asinh", "asin", "acosh", "acos"
        NumberOfOperands = 1
    Case "plus", "minus", "mult", "div", "rem", "pow", "less", "or", "and", "lt", "le", "eq", "ge", "gt", "ne", "atan2", "intdiv", "precision", "round", "trunc", "iff"
        NumberOfOperands = 2
    Case "min", "max", "sum", "count", "numberof", "numberofs", "and_n", "or_n", "alldiff"
        'n-ary - TODO implement
    Case "if", "ifs", "implies"
        NumberOfOperands = 3
    Case Else
        MsgBox "Unknown function " & FunctionName
    End Select
End Function

' Converts common Excel functinos to .nl operators
Function ConvertExcelFunctionToNL(FunctionName As String) As String
    FunctionName = LCase(FunctionName)
    Select Case FunctionName
    Case "ln"
        FunctionName = "log"
    Case "+"
        FunctionName = "plus"
    Case "-"
        FunctionName = "minus"
    Case "*"
        FunctionName = "mult"
    Case "/"
        FunctionName = "div"
    Case "mod"
        FunctionName = "rem"
    Case "^"
        FunctionName = "pow"
    Case "<"
        FunctionName = "lt"
    Case "<="
        FunctionName = "le"
    Case "="
        FunctionName = "eq"
    Case ">="
        FunctionName = "ge"
    Case ">"
        FunctionName = "gt"
    Case "<>"
        FunctionName = "ne"
    Case "quotient"
        FunctionName = "intdiv"
    Case "and"
        FunctionName = "and_n"
    Case "or"
        FunctionName = "or_n"
        
    Case "log", "ceiling", "floor", "power"
        MsgBox "Not implemented yet: " & FunctionName
    End Select
    ConvertExcelFunctionToNL = FunctionName
End Function

' Determines the precedence of arithmetic operators
Function Precedence(tkn As String) As Integer
    Select Case tkn
    Case "plus", "minus"
        Precedence = 2
    Case "mult", "div"
        Precedence = 3
    Case "pow"
        Precedence = 4
    Case "neg"
        Precedence = 5
    Case Else
        Precedence = -1
    End Select
End Function

' Checks the precedence of two operators to determine if the current operator on the stack should be popped
Function CheckPrecedence(tkn1 As String, tkn2 As String) As Boolean
    ' Either tkn1 is left-associative and its precedence is less than or equal to that of tkn2
    If OperatorIsLeftAssociative(tkn1) Then
        If Precedence(tkn1) <= Precedence(tkn2) Then
            CheckPrecedence = True
        Else
            CheckPrecedence = False
        End If
    ' Or tkn1 has precedence less than that of tkn2
    Else
        If Precedence(tkn1) < Precedence(tkn2) Then
            CheckPrecedence = True
        Else
            CheckPrecedence = False
        End If
    End If
End Function

' Determines the left-associativity of arithmetic operators
Function OperatorIsLeftAssociative(tkn As String) As Boolean
    Select Case tkn
    Case "plus", "minus", "mult", "div"
        OperatorIsLeftAssociative = True
    Case "pow"
        OperatorIsLeftAssociative = False
    Case Else
        MsgBox "unknown associativity: " & tkn
    End Select
End Function

' Pops an operator from the operator stack along with the corresponding number of operands.
Sub PopOperator(Operators As StringStack, Operands As ExpressionTreeStack)
    ' Pop the operator and create a new ExpressionTree
    Dim Operator As String
    Operator = Operators.Pop()
    Dim NewTree As ExpressionTree
    Set NewTree = CreateTree(Operator, ExpressionTreeOperator)
    
    ' Pop the required number of operands from the operand stack and set as children of the new operator tree
    Dim NumToPop As Integer, i As Integer, Tree As ExpressionTree
    NumToPop = NumberOfOperands(Operator)
    For i = NumToPop To 1 Step -1
        Set Tree = Operands.Pop
        NewTree.SetChild i, Tree
    Next i
    
    ' Add the new tree to the operands stack
    Operands.Push NewTree
End Sub

' Check whether a token on the operator stack is a function operator (vs. an arithmetic operator)
Function IsFunctionOperator(tkn As String) As Boolean
    Select Case tkn
    Case "plus", "minus", "mult", "div", "pow", "neg"
        IsFunctionOperator = False
    Case Else
        IsFunctionOperator = True
    End Select
End Function

' Negates a tree by adding a 'neg' node to the root
Sub AddNegToTree(Tree As ExpressionTree)
    Dim NewTree As ExpressionTree
    Set NewTree = CreateTree("neg", ExpressionTreeOperator)
    
    Tree.NodeText = right(Tree.NodeText, Len(Tree.NodeText) - 1)
    NewTree.SetChild 1, Tree
    
    Set Tree = NewTree
End Sub

' Formats an expression tree node's text as .nl output
Function FormatNL(NodeText As String, NodeType As ExpressionTreeNodeType) As String
    Select Case NodeType
    Case ExpressionTreeVariable
        FormatNL = "v" & VariableMap(NodeText)
    Case ExpressionTreeNumber
        FormatNL = "n" & NodeText
    Case ExpressionTreeOperator
        FormatNL = "o" & CStr(ConvertOperatorToNLCode(NodeText))
    End Select
End Function

' Converts an operator string to .nl code
Function ConvertOperatorToNLCode(FunctionName As String) As Integer
    Select Case FunctionName
    Case "plus"
        ConvertOperatorToNLCode = 0
    Case "minus"
        ConvertOperatorToNLCode = 1
    Case "mult"
        ConvertOperatorToNLCode = 2
    Case "div"
        ConvertOperatorToNLCode = 3
    Case "rem"
        ConvertOperatorToNLCode = 4
    Case "pow"
        ConvertOperatorToNLCode = 5
    Case "less"
        ConvertOperatorToNLCode = 6
    Case "min"
        ConvertOperatorToNLCode = 11
    Case "max"
        ConvertOperatorToNLCode = 12
    Case "floor"
        ConvertOperatorToNLCode = 13
    Case "ceil"
        ConvertOperatorToNLCode = 14
    Case "abs"
        ConvertOperatorToNLCode = 15
    Case "neg"
        ConvertOperatorToNLCode = 16
    Case "or"
        ConvertOperatorToNLCode = 20
    Case "and"
        ConvertOperatorToNLCode = 21
    Case "lt"
        ConvertOperatorToNLCode = 22
    Case "le"
        ConvertOperatorToNLCode = 23
    Case "eq"
        ConvertOperatorToNLCode = 24
    Case "ge"
        ConvertOperatorToNLCode = 28
    Case "gt"
        ConvertOperatorToNLCode = 29
    Case "ne"
        ConvertOperatorToNLCode = 30
    Case "if"
        ConvertOperatorToNLCode = 35
    Case "not"
        ConvertOperatorToNLCode = 34
    Case "tanh"
        ConvertOperatorToNLCode = 37
    Case "tan"
        ConvertOperatorToNLCode = 38
    Case "sqrt"
        ConvertOperatorToNLCode = 39
    Case "sinh"
        ConvertOperatorToNLCode = 40
    Case "sin"
        ConvertOperatorToNLCode = 41
    Case "log10"
        ConvertOperatorToNLCode = 42
    Case "log"
        ConvertOperatorToNLCode = 43
    Case "exp"
        ConvertOperatorToNLCode = 44
    Case "cosh"
        ConvertOperatorToNLCode = 45
    Case "cos"
        ConvertOperatorToNLCode = 46
    Case "atanh"
        ConvertOperatorToNLCode = 47
    Case "atan2"
        ConvertOperatorToNLCode = 48
    Case "atan"
        ConvertOperatorToNLCode = 49
    Case "asinh"
        ConvertOperatorToNLCode = 50
    Case "asin"
        ConvertOperatorToNLCode = 51
    Case "acosh"
        ConvertOperatorToNLCode = 52
    Case "acos"
        ConvertOperatorToNLCode = 53
    Case "sum"
        ConvertOperatorToNLCode = 54
    Case "intdiv"
        ConvertOperatorToNLCode = 55
    Case "precision"
        ConvertOperatorToNLCode = 56
    Case "round"
        ConvertOperatorToNLCode = 57
    Case "trunc"
        ConvertOperatorToNLCode = 58
    Case "count"
        ConvertOperatorToNLCode = 59
    Case "numberof"
        ConvertOperatorToNLCode = 60
    Case "numberofs"
        ConvertOperatorToNLCode = 61
    Case "ifs"
        ConvertOperatorToNLCode = 65
    Case "and_n"
        ConvertOperatorToNLCode = 70
    Case "or_n"
        ConvertOperatorToNLCode = 71
    Case "implies"
        ConvertOperatorToNLCode = 72
    Case "iff"
        ConvertOperatorToNLCode = 73
    Case "alldiff"
        ConvertOperatorToNLCode = 74
    End Select
End Function

' Converts an objective sense to .nl code
Function ConvertObjectiveSenseToNL(ObjectiveSense As ObjectiveSenseType) As Integer
    Select Case ObjectiveSense
    Case ObjectiveSenseType.MaximiseObjective
        ConvertObjectiveSenseToNL = 1
    Case ObjectiveSenseType.MinimiseObjective
        ConvertObjectiveSenseToNL = 0
    Case Else
        MsgBox "Objective sense not supported: " & ObjectiveSense
    End Select
End Function

' Converts RelationConsts enum to .nl code.
Sub ConvertConstraintToNL(Relation As RelationConsts, BoundType As Integer, Comment As String)
    Select Case Relation
        Case RelationConsts.RelationLE ' Upper Bound on LHS
            BoundType = 1
            Comment = " <= "
        Case RelationConsts.RelationEQ ' Equality
            BoundType = 4
            Comment = " == "
        Case RelationConsts.RelationGE ' Upper Bound on RHS
            BoundType = 2
            Comment = " >= "
    End Select
End Sub

