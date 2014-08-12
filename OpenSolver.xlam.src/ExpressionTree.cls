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

' This class stores an expression tree representing an excel formulae.
' Each ExpressionTree object actually represents a node in the tree, and the tree is built by setting the Children nodes

Dim Children() As ExpressionTree                    ' The children of the current node
Dim NumChildren As Integer                          ' The number of child nodes
Public NodeText As String                           ' The node text - a numeric value, a variable name, or a function name
Public NodeType As ExpressionTreeNodeType           ' The node type - number, variable, or operator

Public NodeLinearity As ExpressionTreeNodeLinearity ' The node linearity - constant, linear or non-linear

Public LinearCoefficient As Double                  ' Track linear coefficient on node when collating linear tree
Public LinearValue As Double                        ' Track linear value of constant node when collating linear tree

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

' Getter method for children
Public Function GetChild(index As Integer) As ExpressionTree
    Dim obj As ExpressionTree
    Set obj = Children(index)
    Set GetChild = obj
End Function

' Setter method for children
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

Public Property Get GetNumChildren() As Integer
    GetNumChildren = NumChildren
End Property

' Clear all children from the node
Public Sub ClearChildren()
    NumChildren = 0
    ReDim Children(0)
End Sub

' Set the defaults for the node on creation
Private Sub Class_Initialize()
    NumChildren = 0
    LinearCoefficient = 0
    NodeLinearity = ExpressionTreeNonLinear
End Sub

' Test is the tree is empty
Public Function IsEmpty() As Boolean
    IsEmpty = NodeText = "0"
End Function


' Negate the tree
Public Function Negate() As ExpressionTree
    If NodeText <> "0" Then
        Dim NewTree As ExpressionTree
        Set NewTree = CreateTree("neg", ExpressionTreeOperator)
        
        NewTree.NodeLinearity = NodeLinearity
        NewTree.SetChild 1, Me
        
        Set Negate = NewTree
    Else
        Set Negate = Me
    End If
End Function

' Display the entire tree as a string for debugging
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

' Display the linearity of a node for debugging
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

' Converts the tree to .nl expression tree format
Public Function ConvertToNL() As String
    ' Construct comment
    Dim Comment As String, i As Integer
    For i = 1 To CommentIndent
        Comment = Comment + " "
    Next i
    Comment = Comment + NodeText + " " + DisplayLinearity()
    
    ' Add current node to output
    Dim output As String
    AddNewLine output, FormatNL(NodeText, NodeType), Comment
    
    ' If n-ary operator, we need the number of operators on a newline
    If IsNAry(NodeText) Then
        CommentIndent = CommentIndent + 2
        Comment = ""
        For i = 1 To CommentIndent
            Comment = Comment + " "
        Next i
        AddNewLine output, CStr(NumChildren), Comment & NumChildren & " arguments"
        CommentIndent = CommentIndent - 2
    End If
    
    ' Add children nodes to output
    CommentIndent = CommentIndent + 4
    For i = 1 To NumChildren
        output = output + Children(i).ConvertToNL
    Next i
    CommentIndent = CommentIndent - 4
    
    ConvertToNL = output
End Function

' Extract all variables from the ExpressionTree into a LinearConstraintNL
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

' Traverse the tree and mark all nodes as Constant, Linear or Non-linear
Sub MarkLinearity()
    ' First process all children
    Dim i As Integer
    For i = 1 To NumChildren
        Children(i).MarkLinearity
    Next i

    Select Case NodeType
    ' A constant node is constant
    Case ExpressionTreeNumber
        NodeLinearity = ExpressionTreeConstant
    
    ' A variable is linear
    Case ExpressionTreeVariable
        NodeLinearity = ExpressionTreeLinear
        
    Case ExpressionTreeOperator
        Select Case NodeText
        Case "plus", "minus"
            ' Non-linear if either child is non-linear
            If Children(1).NodeLinearity = ExpressionTreeNonLinear Or Children(2).NodeLinearity = ExpressionTreeNonLinear Then
                NodeLinearity = ExpressionTreeNonLinear
            ' Otherwise linear if either child is linear
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
            ' Otherwise linear if either child is linear
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
            ' Otherwise non-linear if linear variable raised to power other than one
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

Public Sub PullUpConstants()
    ' First pull all constants in children
    Dim i As Integer
    For i = 1 To NumChildren
        Children(i).PullUpConstants
    Next i

    If NodeLinearity = ExpressionTreeConstant And NodeType = ExpressionTreeOperator Then
        EvaluateConstantOperator
    End If
End Sub

' Evaluates a constant operator node and replaces node and all children with a numeric constant
Public Sub EvaluateConstantOperator()
    Dim i As Integer, Values As New Collection
    For i = 1 To NumChildren
        Values.Add CDbl(Children(i).NodeText)
    Next i
    
    Dim NewValue As Double
    Select Case NodeText
    Case "plus"
        NewValue = Values(1) + Values(2)
    Case "minus"
        NewValue = Values(1) - Values(2)
    Case "mult"
        NewValue = Values(1) * Values(2)
    Case "div"
        NewValue = Values(1) / Values(2)
    Case "pow"
        NewValue = Values(1) ^ Values(2)
    Case "neg"
        NewValue = -Values(1)
    Case "atan"
        NewValue = Atn(Values(1))
    Case Else
        NewValue = 0
        MsgBox "Constant evaluation of " & NodeText & " not implemented yet in collation of linear terms"
    End Select
        
    ' Replace subtree with a constant node
    NodeType = ExpressionTreeNumber
    NodeText = NewValue
    ClearChildren
End Sub

' Traverses the tree and prunes all linear or constant parts from the tree
Public Function PruneLinearTrees(LinearTrees As Collection, Optional NegativeTree As Boolean = False) As Boolean
    Dim CanPrune As Boolean, i As Integer, NewTree As ExpressionTree
    CanPrune = False
    ' If the root node is linear or constant we can prune
    If NodeLinearity <> ExpressionTreeNonLinear Then
       CanPrune = True
    ' Otherwise if the root is a "plus" or "minus", it might have a linear part as one of the children
    ElseIf NodeText = "plus" Or NodeText = "minus" Then
       'Prune children individually and set to n0 nodes
        Dim PrunedFirst As Boolean, PrunedSecond As Boolean
        
        PrunedFirst = Children(1).PruneLinearTrees(LinearTrees)
        If NodeText = "minus" Then
            PrunedSecond = Children(2).PruneLinearTrees(LinearTrees, True)
        Else
            PrunedSecond = Children(2).PruneLinearTrees(LinearTrees)
        End If
        
        ' If first tree pruned, then set this node to the second child
        If PrunedFirst Then
            ' Negate the tree if it was a minus node before setting
            If NodeText = "minus" Then
                SetChild 2, Children(2).Negate
            End If
            SetTreeToChild 2
            
        ' If second node pruned, then set this node to the first child
        ElseIf PrunedSecond Then
            SetTreeToChild 1
        End If
    End If
    
    If CanPrune Then
        ' Copy tree
        Set NewTree = Me.Clone
        
        ' If tree is being pruned from a "minus", we need to negate it
        If NegativeTree Then
            Set NewTree = NewTree.Negate
        End If

        ' Add tree to linear collection
        LinearTrees.Add NewTree
        
        ' Reset current tree to n0
        NodeText = 0
        NodeType = ExpressionTreeNumber
        NodeLinearity = ExpressionTreeConstant
        ClearChildren
        
        PruneLinearTrees = True
    Else
        PruneLinearTrees = False
    End If
        
End Function

' Sets an ExpressionTree to be one of its children
Sub SetTreeToChild(index As Integer)
    ' We can't clear all children of the current tree while referring to one of it's children
    ' We therefore need to make a copy of the child tree before we clear the tree and set it to the child
    Dim NewTree As ExpressionTree
    Set NewTree = Children(index).Clone
    SetExpressionTree NewTree
End Sub

' Creates a copy of an ExpressionTree
Public Function Clone() As ExpressionTree
    Dim NewTree As ExpressionTree
    Set NewTree = CreateTree(NodeText, NodeType)
    NewTree.NodeLinearity = NodeLinearity
    NewTree.LinearCoefficient = LinearCoefficient
    NewTree.LinearValue = LinearValue
    
    Dim i As Integer
    For i = 1 To NumChildren
        NewTree.SetChild i, Children(i)
    Next i
    
    Set Clone = NewTree
End Function

' Sets the current ExpressionTree to be the same as another
Sub SetExpressionTree(NewTree As ExpressionTree)
    NodeText = NewTree.NodeText
    NodeType = NewTree.NodeType
    NodeLinearity = NewTree.NodeLinearity
    LinearCoefficient = NewTree.LinearCoefficient
    LinearValue = NewTree.LinearValue
    
    ClearChildren
    
    Dim i As Integer
    For i = 1 To NewTree.GetNumChildren
        SetChild i, NewTree.GetChild(i)
    Next i
End Sub

' Converts a linear tree into a LinearConstraintNL and Double
Public Sub ConvertLinearTreeToConstraint(constraint As LinearConstraintNL, constant As Double)
    ' Remove everything but "plus", variables, and constants
    CollateLinearTerms
    
    ' Add linear terms into constraint/constant
    AddLinearTreeToConstraint constraint, constant
End Sub

' Traverses a linear tree and replaces all operators other than "plus"
' We want a tree with constants, variables, and "plus" nodes only
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
            EvaluateConstantOperator
            LinearValue = CDbl(NodeText)
        End Select
        
    ' A linear node - we need to update the coefficient on the linear term
    Case ExpressionTreeLinear
        Select Case NodeType
        ' Variable - Intitialise coefficient at 1 if hasn't been already
        Case ExpressionTreeVariable
            If LinearCoefficient = 0 Then
                LinearCoefficient = 1
            End If
        
        ' Linear operator node - we need to collate the constant and variable terms into a single variable node if possible
        Case ExpressionTreeOperator
            Select Case NodeText
            
            ' Do nothing on a plus node. This is just to escape the Else case
            Case "plus"
            
            ' If we have a minus node, convert it to a plus and flip the sign on the second operator
            Case "minus"
                Children(2).CollateLinearNegative
                NodeText = "plus"
                
            ' Negate the coefficient on the variable node and remove the "neg" node
            Case "neg"
                Children(1).MultiplyByConstant -1
                SetTreeToChild 1
                    
            ' There is a constant node and a linear node attached to this mult.
            ' We multiply the constant through the linear node (which may be a "plus" tree)
            Case "mult"
                If Children(1).NodeType = ExpressionTreeConstant Then
                    Children(2).MultiplyByConstant Children(1).LinearValue
                    SetTreeToChild 2
                Else ' Child 2 is the constant node
                    Children(1).MultiplyByConstant Children(2).LinearValue
                    SetTreeToChild 1
                End If
            
            ' First node is variable
            Case "div"
                Children(1).MultiplyByConstant 1 / Children(2).LinearValue
                SetTreeToChild 1
                
            ' Second node must be 1 (as we know the expression is linear)
            Case "pow"
                SetTreeToChild 1
    
            Case Else
                MsgBox "Unknown linear operator " & NodeText
            End Select
        End Select
    End Select
    
    Exit Sub
    
End Sub

' Multiplies all nodes in the linear tree by a constant
Public Sub MultiplyByConstant(constant As Double)
    Select Case NodeType
    Case ExpressionTreeConstant
        LinearValue = LinearValue * constant
        NodeText = LinearValue
    
    Case ExpressionTreeVariable
        LinearCoefficient = LinearCoefficient * constant
        
    Case ExpressionTreeOperator ' plus
        Children(1).MultiplyByConstant constant
        Children(2).MultiplyByConstant constant
        
    End Select
End Sub

' Flips the sign on variable or constant
Public Sub CollateLinearNegative()
    If NodeType = ExpressionTreeVariable Then
        LinearCoefficient = -LinearCoefficient
    Else ' NodeType = ExpressionTreeNumber
        If NodeText <> "0" Then
            LinearValue = -CDbl(NodeText)
            NodeText = LinearValue
        End If
    End If
End Sub

' Convert a collated linear tree into a LinearConstraintNL and Double constant
Sub AddLinearTreeToConstraint(constraint As LinearConstraintNL, constant As Double)
    ' By this point, only "plus", variables and constants remain in the tree.
    ' We traverse the tree and add each variable we find to the constraint,
    ' and each constant value to the constant total
    
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
