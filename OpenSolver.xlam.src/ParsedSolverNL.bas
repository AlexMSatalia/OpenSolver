Attribute VB_Name = "ParsedSolverNL"
Option Explicit

Dim m As CModelParsed

Dim problem_name As String

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

Dim VariableMap As Collection
Dim VariableMapRev As Collection
Dim ConstraintMap As Collection
Dim ConstraintMapRev As Collection

    
Function SolveModelParsed_NL(ModelFilePathName As String, model As CModelParsed)
    Set m = model
    
    Dim numActualVars As Integer, numFakeCons As Integer, numActualCons As Integer, numActualEqs As Integer, numActualRanges As Integer
    numActualVars = m.AdjustableCells.Count
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
    n_obj = 1
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
    
    ProcessConstraints
    
    Open ModelFilePathName For Output As #1
    
    ' Write header
    Print #1, MakeHeader()
    
    ' Write C blocks
    
    ' Write O block
    
    ' Write d block
    Print #1, MakeDBlock()
    
    ' Write x block
    Print #1, MakeXBlock()
    
    ' Write r block
    
    ' Write b block
    Print #1, MakeBBlock()
    
    ' Write k block
    
    ' Write J block
    
    ' Write G block
    
    
    Close #1
End Function

Function MakeHeader() As String
    Dim Header As String
    Header = ""
    
    'Line #1
    AddNewLine Header, "g3 1 1 0 # problem " & problem_name
    'Line #2
    AddNewLine Header, " " & n_var & " " & n_con & " " & n_obj & " " & nranges & " " & n_eqn_ & " " & n_lcon & " # vars, constraints, objectives, ranges, eqns"
    'Line #3
    AddNewLine Header, " " & nlc & " " & nlo & " # nonlinear constraints, objectives"
    'Line #4
    AddNewLine Header, " " & nlnc & " " & lnc & " # network constraints: nonlinear, linear"
    'Line #5
    AddNewLine Header, " " & nlvc & " " & nlvo & " " & nlvb & " # nonlinear vars in constraints, objectives, both"
    'Line #6
    AddNewLine Header, " " & nwv_ & " " & nfunc_ & " " & arith & " " & flags & " # linear network variables; functions; arith, flags"
    'Line #7
    AddNewLine Header, " " & nbv & " " & niv & " " & nlvbi & " " & nlvci & " " & nlvoi & " # discrete variables: binary, integer, nonlinear (b,c,o)"
    'Line #8
    AddNewLine Header, " " & nzc & " " & nzo & " # nonzeros in Jacobian, gradients"
    'Line #9
    AddNewLine Header, " " & maxrownamelen_ & " " & maxcolnamelen_ & " # max name lengths: constraints, variables"
    'Line #10
    AddNewLine Header, " " & comb & " " & comc & " " & como & " " & comc1 & " " & como1 & " # common exprs: b,c,o,c1,o1"
    
    ' Strip trailing newline
    MakeHeader = left(Header, Len(Header) - 2)
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
    For i = 1 To m.LHSKeys.Count
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
    For i = 1 To m.Formulae.Count
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

Sub ProcessConstraints()
    ' Split into linear, non-linear and constant
    Dim i As Integer
    For i = 1 To m.Formulae.Count
        ConvertFormulaToExpressionTree m.Formulae(i).strFormulaParsed
    Next i
End Sub


Function MakeDBlock() As String
    Dim Block As String
    Block = ""
    
    AddNewLine Block, "d" & n_con
    
    ' Set duals to zero for all constraints
    Dim i As Integer
    For i = 1 To n_con
        AddNewLine Block, i - 1 & " 0"
    Next i
    
    ' Strip trailing newline
    MakeDBlock = left(Block, Len(Block) - 2)
End Function

Function MakeXBlock() As String
    Dim Block As String
    Block = ""
    
    AddNewLine Block, "x" & n_var
    
    ' Actual variables, apply bounds
    Dim i As Integer, initial As String
    For i = 1 To m.AdjustableCells.Count
        If VarType(m.AdjustableCells(i)) = vbEmpty Then
            initial = 0
        Else
            initial = m.AdjustableCells(i)
        End If
        AddNewLine Block, i - 1 & " " & initial
    Next i
    
    ' TODO: real initial values for formulae vars
    For i = 1 To m.Formulae.Count
        AddNewLine Block, i + m.AdjustableCells.Count - 1 & " 0"
    Next i
    
    ' Strip trailing newline
    MakeXBlock = left(Block, Len(Block) - 2)
End Function

Function MakeBBlock() As String
    Dim Block As String
    Block = ""
    
    AddNewLine Block, "b"
    
    ' Actual variables, apply bounds
    Dim i As Integer, bound As String
    For i = 1 To m.AdjustableCells.Count
        If m.AssumeNonNegative Then
            bound = "2 0"
        Else
            bound = "3"
        End If
        AddNewLine Block, bound
    Next i
    
    ' Fake formulae variables - no bounds
    For i = 1 To m.Formulae.Count
        AddNewLine Block, "3"
    Next i
    
    ' Strip trailing newline
    MakeBBlock = left(Block, Len(Block) - 2)
End Function

Sub AddNewLine(CurText As String, LineText As String)
    CurText = CurText & LineText & vbNewLine
End Sub

Sub ConvertFormulaToExpressionTree(strFormula As String)
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
        Case TokenType.SubExpressionClose
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
    Debug.Print Operands.Peek.Display
    
    Exit Sub
    
Mismatch:
    MsgBox "Mismatched parentheses"
    
End Sub

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
    Case "min", "max", "sum", "count", "numberof", "numberofs", "and", "or", "alldiff"
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
