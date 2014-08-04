VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExpressionTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim Parent As ExpressionTree
Dim Children() As ExpressionTree
Dim NumChildren As Integer
Public NodeText As String
Public NodeType As ExpressionTreeNodeType
Public NodeLinearity As ExpressionTreeNodeLinearity

Public LinearCoefficient As Double ' Track linear coefficient on node when collating linear tree
Public LinearValue As Double ' Track linear value of constant node

Public Enum ExpressionTreeNodeType
    ExpressionTreeNumber = 0
    ExpressionTreeVariable = 1
    ExpressionTreeOperator = 2
End Enum

Public Enum ExpressionTreeNodeLinearity
    ExpressionTreeConstant = 0
    ExpressionTreeLinear = 1
    ExpressionTreeNonLinear = 2
End Enum

Public Function GetChild(index As Integer) As ExpressionTree
    Dim obj As ExpressionTree
    Set obj = Children(i)
    Set GetChild = obj
End Function

Public Sub SetChild(index As Integer, obj As ExpressionTree)
    If NumChildren = 0 Then
        ReDim Preserve Children(index)
        NumChildren = index
    ElseIf index > UBound(Children) Then
        ReDim Preserve Children(index)
        NumChildren = index
    End If
    Set Children(index) = obj
End Sub

Public Sub ClearChildren()
    NumChildren = 0
    ReDim Children(0)
End Sub

Private Sub Class_Initialize()
    NumChildren = 0
    LinearCoefficient = 0
    NodeLinearity = ExpressionTreeNonLinear
End Sub

Public Function Negate() As ExpressionTree
    If NodeText <> "0" Then
        Dim NewTree As ExpressionTree
        Set NewTree = CreateTree("neg", ExpressionTreeOperator)
        
        NewTree.SetChild 1, Me
        
        Set Negate = NewTree
    Else
        Set Negate = Me
    End If
End Function

Public Function Display() As String
    Dim output As String
    If NodeType = ExpressionTreeOperator Then
        output = "("
    Else
        output = ""
    End If
    
    Dim i As Integer
    For i = 1 To NumChildren
        output = output + Children(i).Display
    Next i
    
    output = output + NodeText
    
    If NodeType = ExpressionTreeOperator Then
        output = output + ") "
    Else
        output = output + " "
    End If
    
    Display = output
End Function

Function DisplayLinearity() As String
    Select Case NodeLinearity
    Case ExpressionTreeNonLinear
        DisplayLinearity = "N"
    Case ExpressionTreeLinear
        DisplayLinearity = "L"
    Case ExpressionTreeConstant
        DisplayLinearity = "C"
    End Select
End Function

Public Function ConvertToNL() As String
    
    ' Add comment
    Dim comment As String, i As Integer
    For i = 1 To CommentIndent
       comment = comment + " "
    Next i
    comment = comment + NodeText + " " + DisplayLinearity()
    
    Dim output As String
    AddNewLine output, FormatNL(NodeText, NodeType), comment
    
    CommentIndent = CommentIndent + 4
    For i = 1 To NumChildren
        output = output + Children(i).ConvertToNL
    Next i
    CommentIndent = CommentIndent - 4
    
    ConvertToNL = output
End Function

Sub ExtractVariables(constraint As LinearConstraintNL)
    If NodeType = ExpressionTreeVariable Then
        Dim index As Integer
        index = VariableIndex(NodeText)
        If Not constraint.VariablePresent(index) Then
            constraint.VariablePresent(index) = True
        End If
    End If
    
    Dim i As Integer
    For i = 1 To NumChildren
        Children(i).ExtractVariables constraint
    Next i

End Sub

Sub MarkLinearity()
    ' First process all children
    Dim i As Integer
    For i = 1 To NumChildren
        Children(i).MarkLinearity
    Next i

    Select Case NodeType
    Case ExpressionTreeNumber
        NodeLinearity = ExpressionTreeConstant
        
    Case ExpressionTreeVariable
        NodeLinearity = ExpressionTreeLinear
        
    Case ExpressionTreeOperator
        Select Case NodeText
        Case "plus", "minus"
            ' Non-linear if either child is non-linear
            If Children(1).NodeLinearity = ExpressionTreeNonLinear Or Children(2).NodeLinearity = ExpressionTreeNonLinear Then
                NodeLinearity = ExpressionTreeNonLinear
            ' Linear if either child is linear
            ElseIf Children(1).NodeLinearity = ExpressionTreeLinear Or Children(2).NodeLinearity = ExpressionTreeLinear Then
                NodeLinearity = ExpressionTreeLinear
            ' Otherwise constant
            Else
                NodeLinearity = ExpressionTreeConstant
            End If
            
        Case "mult"
            ' Non-linear if either child is non-linear or both are linear
            If Children(1).NodeLinearity = ExpressionTreeNonLinear Or Children(2).NodeLinearity = ExpressionTreeNonLinear Or _
               (Children(1).NodeLinearity = ExpressionTreeLinear And Children(2).NodeLinearity = ExpressionTreeLinear) Then
                NodeLinearity = ExpressionTreeNonLinear
            ' Linear if either child is linear
            ElseIf Children(1).NodeLinearity = ExpressionTreeLinear Or Children(2).NodeLinearity = ExpressionTreeLinear Then
                NodeLinearity = ExpressionTreeLinear
            ' Otherwise constant
            Else
                NodeLinearity = ExpressionTreeConstant
            End If
            
        Case "div"
            ' If denominator is not constant, expression is non-linear
            If Not Children(2).NodeLinearity = ExpressionTreeConstant Then
                NodeLinearity = ExpressionTreeNonLinear
            ' Otherwise denominator is constant and expression is same linearity as numerator
            Else
                NodeLinearity = Children(1).NodeLinearity
            End If
        
        Case "pow"
            ' If exponent is not constant, then non-linear
            If Not Children(2).NodeLinearity = ExpressionTreeConstant Then
               NodeLinearity = ExpressionTreeNonLinear
            ' Non-linear if linear variable raised to power other than one
            ElseIf Children(1).NodeLinearity = ExpressionTreeLinear And Children(2).NodeText <> 1 Then
               NodeLinearity = ExpressionTreeNonLinear
            ' Otherwise expression is same linearity as base
            Else
               NodeLinearity = Children(1).NodeLinearity
            End If
        
        Case "neg"
            ' Same linearity as child
            NodeLinearity = Children(1).NodeLinearity
            
        Case Else
            ' Other functions
            ' If all arguments const, then node is const
            ' Otherwise non-linear
            ' This is a catch-all, specific functions might need more detail
            For i = 1 To NumChildren
                If Not Children(i).NodeLinearity = ExpressionTreeConstant Then
                    NodeLinearity = ExpressionTreeNonLinear
                    Exit Sub
                End If
            Next i
            NodeLinearity = ExpressionTreeConstant
            
        End Select
    End Select
End Sub

Public Sub PruneLinearTrees(LinearTrees As Collection, Optional RootNode As Boolean = False)
    Dim CanPrune As Boolean
    CanPrune = False
    If NodeLinearity <> ExpressionTreeNonLinear Then
       CanPrune = True
    ElseIf NodeText = "sum" Or NodeText = "minus" Then
       'TODO prune children individually and set to n0 nodes
    Else
    End If
    
    If CanPrune Then
        ' Copy tree
        Dim NewTree As ExpressionTree
        Set NewTree = CreateTree(NodeText, NodeType)
        NewTree.NodeLinearity = NodeLinearity
        Dim i As Integer
        For i = 1 To NumChildren
            NewTree.SetChild i, Children(i)
        Next i

        ' Add tree to linear collection
        LinearTrees.Add NewTree
        
        ' Reset current tree to n0
        NodeText = 0
        NodeType = ExpressionTreeNumber
        NodeLinearity = ExpressionTreeConstant
        ClearChildren
    End If
        
End Sub

Public Sub ConvertLinearTreeToConstraint(constraint As LinearConstraintNL, constant As Double)
    ' Remove everything but +, -, variables, and constants
    CollateLinearTerms
    
    ' Add linear terms into constraint/constant
    AddLinearTreeToConstraint constraint, constant
End Sub

Sub CollateLinearTerms()

    ' First collate all children
    Dim i As Integer
    For i = 1 To NumChildren
        Children(i).CollateLinearTerms
    Next i

    Select Case NodeLinearity
    ' If the node if constant, we want to evaluate it
    Case ExpressionTreeConstant
        Select Case NodeType
        ' Constant expression on a number node - we read the number
        Case ExpressionTreeNumber
            LinearValue = CDbl(NodeText)
        
        ' Constant expression under operator - we can evaluate it and discard all children
        Case ExpressionTreeOperator
            Select Case NodeText
            Case "plus"
                LinearValue = Children(1).LinearValue + Children(2).LinearValue
            Case "minus"
                LinearValue = Children(1).LinearValue - Children(2).LinearValue
            Case "mult"
                LinearValue = Children(1).LinearValue * Children(2).LinearValue
            Case "div"
                LinearValue = Children(1).LinearValue / Children(2).LinearValue
            Case "pow"
                LinearValue = Children(1).LinearValue ^ Children(2).LinearValue
            Case "neg"
                LinearValue = -Children(1).LinearValue
            Case Else
                LinearValue = 0
                MsgBox "Constant evaluation of " & NodeText & " not implemented yet in collation of linear terms"
            End Select
            
            ' Replace subtree with a constant node
            NodeType = ExpressionTreeNumber
            NodeText = LinearValue
            ClearChildren
        End Select
        
    ' A linear node - we need to update the coefficient on the linear term
    Case ExpressionTreeLinear
        Select Case NodeType
        ' Variable - Intitialise coefficient at 1 if hasn't been already
        Case ExpressionTreeVariable
            If LinearCoefficient = 0 Then
                LinearCoefficient = 1
            End If
        
        ' Linear operator node - we need to collate the constant and variable terms into a single variable node
        Case ExpressionTreeOperator
            ' Make sure there is a constant and variable to collect
            If (Children(1).NodeType = ExpressionTreeNumber And Children(2).NodeType = ExpressionTreeVariable) Or _
               (Children(1).NodeType = ExpressionTreeVariable And Children(2).NodeType = ExpressionTreeNumber) Then
                Select Case NodeText

                ' Need to detect which node is variable
                Case "mult"
                    If Children(1).NodeType = ExpressionTreeVariable Then
                        LinearCoefficient = Children(1).LinearCoefficient * Children(2).LinearValue
                        NodeText = Children(1).NodeText
                    Else
                        LinearCoefficient = Children(2).LinearCoefficient * Children(1).LinearValue
                        NodeText = Children(2).NodeText
                    End If
                
                ' First node is variable
                Case "div"
                    LinearCoefficient = Children(1).LinearCoefficient / Children(2).LinearValue
                    NodeText = Children(1).NodeText
                    
                ' Second node must be 1 (as we know the expression is linear)
                Case "pow"
                    LinearCoefficient = Children(1).LinearCoefficient
                    NodeText = Children(1).NodeText
                
                Case "neg"
                    LinearCoefficient = -Children(1).LinearCoefficient
                    NodeText = Children(1).NodeText
        
                Case Else
                    MsgBox "Unknown linear operator " & NodeText
                End Select
                
                ' Remove subtrees
                NodeType = ExpressionTreeVariable
                ClearChildren
            
            ' If we have a minus node, convert it to a plus and flip the sign on the second operator
            ElseIf NodeText = "minus" Then
                If Children(2).NodeType = ExpressionTreeVariable Then
                    Children(2).LinearCoefficient = -Children(2).LinearCoefficient
                Else ' Children(2).NodeType = ExpressionTreeNumber
                    Children(2).LinearValue = -CDbl(NodeText)
                    Children(2).NodeText = Children(2).LinearValue
                End If
                
                NodeText = "plus"
            End If
        End Select
    End Select
    
    Exit Sub
    
End Sub


Sub AddLinearTreeToConstraint(constraint As LinearConstraintNL, constant As Double)
    ' By this point, only "plus", variables and constants remain in the tree
    Select Case NodeType
    ' Add variable coefficient to constraint
    Case ExpressionTreeVariable
        Dim VarIndex As Integer
        VarIndex = VariableIndex(NodeText)
        constraint.Coefficient(VarIndex) = constraint.Coefficient(VarIndex) + LinearCoefficient
    
    ' Add constant
    Case ExpressionTreeNumber
        constant = constant + LinearValue
    
    ' Recurse on all "plus" nodes
    Case ExpressionTreeOperator
        If NodeText = "plus" Then
            Dim i As Integer
            For i = 1 To NumChildren
                Children(i).AddLinearTreeToConstraint constraint, constant
            Next i
        Else
            MsgBox "Operator that isn't plus at the last stage of the linear tree pruning: " & NodeText
        End If
    End Select
End Sub
