Attribute VB_Name = "ParsedSolverNL"
Option Explicit

Dim m As CModelParsed

Dim problem_name As String

Public CommentIndent As Integer     ' Tracks the level of indenting in comments on nl output
Public Const CommentSpacing = 20    ' The column number at which nl comments begin

' ==========================================================================
' ASL variables
' These are variables used in the .NL file that are also used by the AMPL Solver Library
' We use the same names as the ASL for consistency
' ASL header definitions available at http://www.netlib.org/ampl/solvers/asl.h

Dim n_var As Integer    ' Number of variables
Dim n_con As Integer    ' Number of constraints
Dim n_obj As Integer    ' Number of objectives
Dim nranges As Integer  ' Number of range constraints
Dim n_eqn_ As Integer   ' Number of equality constraints
Dim n_lcon As Integer   ' Number of logical constraints

Dim nlc As Integer  ' Number of non-linear constraints
Dim nlo As Integer  ' Number of non-linear objectives

Dim nlnc As Integer ' Number of non-linear network constraints
Dim lnc As Integer  ' Number of linear network constraints

Dim nlvc As Integer ' Number of variables appearing non-linearly in constraints
Dim nlvo As Integer ' Number of variables appearing non-linearly in objectives
Dim nlvb As Integer ' Number of variables appearing non-linearly in both constraints and objectives

Dim nwv_ As Integer     ' Number of linear network variables
Dim nfunc_ As Integer   ' Number of user-defined functions
Dim arith As Integer    ' Not sure what this does, keep at 0. This flag may indicate whether little- or big-endian arithmetic was used to write the file.
Dim flags As Integer    ' 1 = want output suffixes

Dim nbv As Integer      ' Number of binary variables
Dim niv As Integer      ' Number of other integer variables
Dim nlvbi As Integer    ' Number of integer variables appearing non-linearly in both constraints and objectives
Dim nlvci As Integer    ' Number of integer variables appearing non-linearly just in constraints
Dim nlvoi As Integer    ' number of integer variables appearing non-linearly just in objectives

Dim nzc As Integer  ' Number of non-zeros in the Jacobian matrix
Dim nzo As Integer  ' Number of non-zeros in objective gradients

Dim maxrownamelen_ As Integer ' Length of longest constraint name
Dim maxcolnamelen_ As Integer ' Length of longest variable name

' Common expressions (from "defined" vars) - whatever that means!? We just set them to zero as we don't use defined vars
Dim comb As Integer
Dim comc As Integer
Dim como As Integer
Dim comc1 As Integer
Dim como1 As Integer

' End ASL variables
' ===========================================================================

Public VariableMap As Collection
Dim VariableMapRev As Collection
Dim ConstraintMap As Collection
Dim ConstraintMapRev As Collection

Dim NonLinearConstraintTrees As Collection
Dim NonLinearObjectiveTrees As Collection

Dim NonLinearVars() As Boolean

Dim LinearConstraints As Collection
Dim LinearObjectives As Collection

Dim ObjectiveCells As Collection
Dim ObjectiveSenses As Collection

Dim numActualVars As Integer
Dim numFakeVars As Integer
Dim numFakeCons As Integer
Dim numActualCons As Integer
Dim numActualEqs As Integer
Dim numActualRanges As Integer

Dim NonZeroConstraintCount() As Integer

    
Function SolveModelParsed_NL(ModelFilePathName As String, model As CModelParsed)
    Set m = model
    
    numActualVars = m.AdjustableCells.Count
    numFakeVars = m.Formulae.Count
    numFakeCons = m.Formulae.Count
    numActualCons = m.LHSKeys.Count
    
    Dim ActualRangeConstraints As New Collection, ActualEqConstraints As New Collection
    Dim i As Integer
    For i = 1 To numActualCons
        If m.Rels(i) = RelationConsts.RelationEQ Then
            ActualEqConstraints.Add i
        Else
            ActualRangeConstraints.Add i
        End If
    Next i
    
    numActualEqs = ActualEqConstraints.Count
    numActualRanges = ActualRangeConstraints.Count
    
    ' Model statistics for line #1
    problem_name = "'Sheet=" + m.SolverModelSheet.Name + "'"
    
    ' Model statistics for line #2
    n_var = numActualVars + numFakeCons
    n_con = numActualCons + numFakeCons
    n_obj = 0
    nranges = numActualRanges
    n_eqn_ = numActualEqs + numFakeCons
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
    
    MakeVariableMap
    MakeConstraintMap
    
    OutputColFile
    OutputRowFile
    
    ProcessFormulae
    ProcessObjective
    
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
    Exit Function
    
ErrHandler:
    Close #1
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

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
    
    ' Strip trailing newline
    MakeHeader = StripTrailingNewline(Header)
End Function

Sub MakeVariableMap()
    maxcolnamelen_ = 0
    Set VariableMap = New Collection
    Set VariableMapRev = New Collection
    
    Dim index As Integer
    index = 0
    
    ' Actual variables
    Dim c As Range, cellName As String
    For Each c In m.AdjustableCells
        cellName = ConvertCellToStandardName(c)
        
        ' Update variable maps
        VariableMap.Add CStr(index), cellName
        VariableMapRev.Add cellName, CStr(index)
        
        ' Update max length
        If Len(cellName) > maxcolnamelen_ Then
           maxcolnamelen_ = Len(cellName)
        End If
        index = index + 1
    Next c
    
    ' Formulae variables
    Dim i As Integer
    For i = 1 To m.Formulae.Count
        cellName = m.Formulae(i).strAddress
        
        ' Update variable maps
        VariableMap.Add CStr(index), cellName
        VariableMapRev.Add cellName, CStr(index)
        
        ' Update max length
        If Len(cellName) > maxcolnamelen_ Then
           maxcolnamelen_ = Len(cellName)
        End If
        index = index + 1
    Next i
End Sub

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

Sub MakeConstraintMap()
    maxrownamelen_ = 0
    Set ConstraintMap = New Collection
    Set ConstraintMapRev = New Collection
    
    Dim index As Integer
    index = 0
    
    ' Actual constraints
    Dim i As Integer, cellName As String
    For i = 1 To numActualCons
        cellName = "c" & i & "_" & m.LHSKeys(i)
    
        ' Update constraint maps
        ConstraintMap.Add index, cellName
        ConstraintMapRev.Add cellName, CStr(index)
        
        ' Update max length
        If Len(cellName) > maxrownamelen_ Then
           maxrownamelen_ = Len(cellName)
        End If
        index = index + 1
    Next i
    
    ' Formulae variables
    For i = 1 To numFakeCons
        cellName = "f" & i & "_" & m.Formulae(i).strAddress
        
        ' Update variable maps
        ConstraintMap.Add index, cellName
        ConstraintMapRev.Add cellName, CStr(index)
        
        ' Update max length
        If Len(cellName) > maxrownamelen_ Then
           maxrownamelen_ = Len(cellName)
        End If
        index = index + 1
    Next i
End Sub

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

Sub ProcessFormulae()
    
    Set NonLinearConstraintTrees = New Collection
    Set LinearConstraints = New Collection
    
    ReDim NonLinearVars(n_var)
    
    
    Dim i As Integer
    For i = 1 To numActualCons
        ProcessSingleFormula m.RHSKeys(i), m.LHSKeys(i)
    Next i
    
    For i = 1 To numFakeCons
        ProcessSingleFormula m.Formulae(i).strFormulaParsed, m.Formulae(i).strAddress
    Next i
    
    ' Process linear vars for jacobian counts
    ReDim NonZeroConstraintCount(n_var)
    Dim j As Integer
    For i = 1 To LinearConstraints.Count
        For j = 1 To LinearConstraints(i).Count
            If LinearConstraints(i).VariablePresent(j) And LinearConstraints(i).Coefficient(j) <> 0 Then
                NonZeroConstraintCount(j) = NonZeroConstraintCount(j) + 1
                nzc = nzc + 1
            End If
        Next j
    Next i
    
End Sub

Sub ProcessSingleFormula(RHSExpression As String, LHSVariable As String)
    Dim Tree As ExpressionTree
    Set Tree = ConvertFormulaToExpressionTree(RHSExpression)
    Debug.Print Tree.Display
    
    
    Dim constraint As LinearConstraintNL
    Set constraint = New LinearConstraintNL
    constraint.Count = n_var
    
    Tree.ExtractVariables constraint
    
    Dim j As Integer
    For j = 1 To constraint.Count
        If constraint.VariablePresent(j) And Not NonLinearVars(j) Then
            NonLinearVars(j) = True
            nlvc = nlvc + 1
        End If
    Next j

    ' Remove linear terms
    Tree.MarkLinearity
    ' Set header information
    
    NonLinearConstraintTrees.Add Tree
    
    constraint.VariablePresent(VariableMap(LHSVariable) + 1) = True
    constraint.Coefficient(VariableMap(LHSVariable) + 1) = -1
    
    LinearConstraints.Add constraint
    
    nlc = nlc + 1
End Sub

Sub ProcessObjective()
' Currently just one objective
' Objective has a single linear term - the objective variable

    Set NonLinearObjectiveTrees = New Collection
    Set ObjectiveSenses = New Collection
    Set ObjectiveCells = New Collection
    Set LinearObjectives = New Collection

    ObjectiveCells.Add ConvertCellToStandardName(m.ObjectiveCell)
    ObjectiveSenses.Add m.ObjectiveSense
    
    ' Objective non-linear constraint tree is empty
    NonLinearObjectiveTrees.Add CreateTree("0", ExpressionTreeNodeType.ExpressionTreeNumber)
    
    ' Objective linear constraint
    Dim Objective As New LinearConstraintNL
    Objective.Count = n_var
    
    Objective.VariablePresent(VariableMap(ObjectiveCells(1)) + 1) = True
    Objective.Coefficient(VariableMap(ObjectiveCells(1)) + 1) = 1
    
    LinearObjectives.Add Objective
    ' Track non-zero count in linear objective
    nzo = nzo + 1
    
    ' Adjust objective count
    n_obj = n_obj + 1
End Sub

Function MakeCBlocks() As String
    Dim Block As String
    Block = ""
    
    Dim i As Integer
    For i = 1 To n_con
        ' Add block header
        AddNewLine Block, "C" & i - 1, "CONSTRAINT NON-LINEAR SECTION " + ConstraintMapRev(CStr(i - 1))
        
        ' Add expression tree
        CommentIndent = 4
        Block = Block + NonLinearConstraintTrees(i).ConvertToNL
    Next i
    
    MakeCBlocks = StripTrailingNewline(Block)
End Function

Function MakeOBlocks() As String
    
    Dim Block As String
    Block = ""
    
    Dim i As Integer
    For i = 1 To n_obj
        AddNewLine Block, "O" & i - 1 & " " & ConvertObjectiveSenseToNL(ObjectiveSenses(i)), "OBJECTIVE NON-LINEAR SECTION " & ObjectiveCells(i)
        
        ' Add expression tree
        CommentIndent = 4
        Block = Block + NonLinearObjectiveTrees(i).ConvertToNL
    Next i
    
    MakeOBlocks = StripTrailingNewline(Block)
End Function

Function MakeDBlock() As String
    Dim Block As String
    Block = ""
    
    AddNewLine Block, "d" & n_con, "INITIAL DUAL GUESS"
    
    ' Set duals to zero for all constraints
    Dim i As Integer
    For i = 1 To n_con
        AddNewLine Block, i - 1 & " 0", "    " & ConstraintMapRev(CStr(i - 1)) & " = " & 0
    Next i
    
    ' Strip trailing newline
    MakeDBlock = StripTrailingNewline(Block)
End Function

Function MakeXBlock() As String
    Dim Block As String
    Block = ""
    
    AddNewLine Block, "x" & n_var, "INITIAL PRIMAL GUESS"
    
    ' Actual variables, apply bounds
    Dim i As Integer, initial As String
    For i = 1 To numActualVars
        If VarType(m.AdjustableCells(i)) = vbEmpty Then
            initial = 0
        Else
            initial = m.AdjustableCells(i)
        End If
        AddNewLine Block, i - 1 & " " & initial, "    " & VariableMapRev(CStr(i - 1)) & " = " & initial
    Next i
    
    ' Initial values for formulae vars
    For i = 1 To numFakeVars
        initial = CStr(m.Formulae(i).initialValue)
        AddNewLine Block, i + numActualVars - 1 & " " & initial, "    " & VariableMapRev(CStr(i - 1 + numActualVars)) & " = " & initial
    Next i
    
    ' Strip trailing newline
    MakeXBlock = StripTrailingNewline(Block)
End Function

Function MakeRBlock() As String
    Dim Block As String
    Block = ""
    
    AddNewLine Block, "r", "CONSTRAINT BOUNDS"
    
    ' Actual constraints - apply bounds
    Dim i As Integer, BoundType As Integer, comment As String
    For i = 1 To numActualCons
        ConvertConstraintToNL m.Rels(i), BoundType, comment
        AddNewLine Block, BoundType & " " & 0, "    " & ConstraintMapRev(CStr(i - 1)) & comment & 0
    Next i
    
    ' Fake formulae constraints - must equal 0
    For i = 1 To numFakeCons
        ConvertConstraintToNL RelationConsts.RelationEQ, BoundType, comment
        AddNewLine Block, BoundType & " " & 0, "    " & ConstraintMapRev(CStr(i - 1 + numActualCons)) & comment & 0
    Next i
    
    ' Strip trailing newline
    MakeRBlock = StripTrailingNewline(Block)
End Function

Function MakeBBlock() As String
    Dim Block As String
    Block = ""
    
    AddNewLine Block, "b", "VARIABLE BOUNDS"
    
    ' Actual variables, apply bounds
    Dim i As Integer, bound As String, comment As String
    For i = 1 To numActualVars
        comment = "    " & VariableMapRev(CStr(i - 1))
        If m.AssumeNonNegative Then
            bound = "2 0"
            comment = comment & " >= 0"
        Else
            bound = "3"
            comment = comment & " FREE"
        End If
        AddNewLine Block, bound, comment
    Next i
    
    ' Fake formulae variables - no bounds
    For i = 1 To numFakeVars
        AddNewLine Block, "3", "    " & VariableMapRev(CStr(i - 1 + numActualVars)) & " FREE"
    Next i
    
    ' Strip trailing newline
    MakeBBlock = StripTrailingNewline(Block)
End Function

Function MakeKBlock() As String
    Dim Block As String
    Block = ""

    ' Add header
    AddNewLine Block, "k" & n_var - 1, "NUMBER OF JACOBIAN ENTRIES (CUMULATIVE) FOR FIRST " & n_var - 1 & " VARIABLES"
    
    Dim i As Integer, total As Integer
    total = 0
    For i = 1 To n_var - 1
        total = total + NonZeroConstraintCount(i)
        AddNewLine Block, CStr(total), "    Up to " & VariableMapRev(CStr(i - 1)) & ": " & CStr(total) & " entries in Jacobian"
    Next i
    
    MakeKBlock = StripTrailingNewline(Block)
End Function

Function MakeJBlocks() As String
    Dim Block As String
    Block = ""
    
    Dim i As Integer, ObjectiveVariables As Collection, ObjectiveCoefficients As Collection
    For i = 1 To n_con
        ' Make header
        AddNewLine Block, "J" & i - 1 & " " & LinearConstraints(i).NumPresent, "CONSTRAINT LINEAR SECTION " & ConstraintMapRev(i)
        
        Dim j As Integer
        For j = 1 To LinearConstraints(i).Count
            If LinearConstraints(i).VariablePresent(j) Then
                AddNewLine Block, j - 1 & " " & LinearConstraints(i).Coefficient(j), "    + " & LinearConstraints(i).Coefficient(j) & " * " & VariableMapRev(CStr(j - 1))
            End If
        Next j
    Next i
    
    MakeJBlocks = StripTrailingNewline(Block)
End Function

Function MakeGBlocks() As String
    Dim Block As String
    Block = ""
    
    Dim i As Integer, ObjectiveVariables As Collection, ObjectiveCoefficients As Collection
    For i = 1 To n_obj
        ' Make header
        AddNewLine Block, "G" & i - 1 & " " & LinearObjectives(i).NumPresent, "OBJECTIVE LINEAR SECTION " & ObjectiveCells(i)
        
        Dim j As Integer
        For j = 1 To LinearObjectives(i).Count
            If LinearObjectives(i).VariablePresent(j) Then
                AddNewLine Block, j - 1 & " " & LinearObjectives(i).Coefficient(j), "    + " & LinearObjectives(i).Coefficient(j) & " * " & VariableMapRev(CStr(j - 1))
            End If
        Next j
    Next i
    
    MakeGBlocks = StripTrailingNewline(Block)
End Function

Sub AddNewLine(CurText As String, LineText As String, Optional CommentText As String = "")
    Dim Space As String
    Space = ""
    
    If CommentText <> "" Then
        Dim j As Integer
        For j = 1 To CommentSpacing - Len(LineText)
           Space = Space + " "
        Next j
        Space = Space + "# "
    End If
    
    CurText = CurText & LineText & Space & CommentText & vbNewLine
End Sub

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

Public Function CreateTree(NodeText As String, NodeType As Long) As ExpressionTree
    Dim obj As ExpressionTree

    Set obj = New ExpressionTree
    obj.NodeText = NodeText
    obj.NodeType = NodeType

    Set CreateTree = obj
End Function

Function NumberOfOperands(FunctionName As String) As Integer
    Select Case FunctionName
    Case "floor", "ceil", "abs", "neg", "not", "tanh", "tan", "sqrt", "sinh", "sin", "log10", "log", "exp", "cosh", "cos", "atanh", "atan", "asinh", "asin", "acosh", "acos"
        NumberOfOperands = 1
    Case "plus", "minus", "mult", "div", "rem", "pow", "less", "or", "and", "lt", "le", "eq", "ge", "gt", "ne", "atan2", "intdiv", "precision", "round", "trunc", "iff"
        NumberOfOperands = 2
    Case "min", "max", "sum", "count", "numberof", "numberofs", "and_n", "or_n", "alldiff"
        'n-ary
    Case "if", "ifs", "implies"
        NumberOfOperands = 3
    Case Else
        MsgBox "Unknown function " & FunctionName
    End Select
End Function

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

Sub PopOperator(Operators As StringStack, Operands As ExpressionTreeStack)
    Dim Operator As String
    Operator = Operators.Pop()
    
    Dim NewTree As ExpressionTree
    Set NewTree = CreateTree(Operator, ExpressionTreeOperator)
    
    Dim NumToPop As Integer
    NumToPop = NumberOfOperands(Operator)
    
    Dim i As Integer, Tree As ExpressionTree
    For i = NumToPop To 1 Step -1
        Set Tree = Operands.Pop
        NewTree.SetChild i, Tree
    Next i
    
    Operands.Push NewTree
End Sub

Function IsFunctionOperator(tkn As String) As Boolean
    Select Case tkn
    Case "plus", "minus", "mult", "div", "pow", "neg"
        IsFunctionOperator = False
    Case Else
        IsFunctionOperator = True
    End Select
End Function

Sub AddNegToTree(Tree As ExpressionTree)
    Dim NewTree As ExpressionTree
    Set NewTree = CreateTree("neg", ExpressionTreeOperator)
    
    Tree.NodeText = right(Tree.NodeText, Len(Tree.NodeText) - 1)
    NewTree.SetChild 1, Tree
    
    Set Tree = NewTree
End Sub

Sub AddLHSToExpressionTree(Tree As ExpressionTree, LHS As String)
' Bring the LHS variable over to the RHS as a minus operation
    Dim NewTree As ExpressionTree
    Set NewTree = CreateTree("minus", ExpressionTreeOperator)
    
    NewTree.SetChild 1, Tree
    Set Tree = NewTree
    
    Set NewTree = CreateTree(LHS, ExpressionTreeVariable)
    Tree.SetChild 2, NewTree
End Sub

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

Sub ConvertConstraintToNL(Relation As RelationConsts, BoundType As Integer, comment As String)
' Converts RelationConsts enum to NL format.
' LE and GE are swapped here because the relation const is defined w.r.t to LHS but the NL deals with the RHS
' so LE becomes >= in the NL file, and GE becomes <=
    Select Case Relation
        Case RelationConsts.RelationLE
            BoundType = 1
            comment = " >= "
        Case RelationConsts.RelationEQ
            BoundType = 4
            comment = " == "
        Case RelationConsts.RelationGE
            BoundType = 2
            comment = " <= "
    End Select
End Sub

