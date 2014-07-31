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

Private Sub Class_Initialize()
    NumChildren = 0
    NodeLinearity = ExpressionTreeNonLinear
End Sub

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
        index = VariableMap(NodeText) + 1
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
