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

Public Enum ExpressionTreeNodeType
    ExpressionTreeNumber = 0
    ExpressionTreeVariable = 1
    ExpressionTreeOperator = 2
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

Public Function ConvertToNL() As String
    
    ' Add comment
    Dim comment As String, i As Integer
    For i = 1 To CommentIndent
       comment = comment + " "
    Next i
    comment = comment + NodeText
    
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
