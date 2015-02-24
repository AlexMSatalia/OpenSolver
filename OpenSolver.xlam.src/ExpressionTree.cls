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
Dim NumChildren As Long                          ' The number of child nodes
Public NodeText As String                           ' The node text - a numeric value, a variable name, or a function name
Public NodeType As ExpressionTreeNodeType           ' The node type - number, variable, or operator

Public NodeLinearity As ExpressionTreeNodeLinearity ' The node linearity - constant, linear or non-linear

Public LinearCoefficient As Double                  ' Track linear coefficient on node when collating linear tree
Public LinearValue As Double                        ' Track linear value of constant node when collating linear tree

Public Enum ExpressionTreeNodeType
    ExpressionTreeNumber = 0
    ExpressionTreeVariable = 1
    ExpressionTreeOperator = 2
    ExpressionTreeString = 3
End Enum

Public Enum ExpressionTreeNodeLinearity
    ExpressionTreeConstant = 0
    ExpressionTreeLinear = 1
    ExpressionTreeNonLinear = 2
End Enum

' Getter method for children
Public Function GetChild(Index As Long) As ExpressionTree
          Dim obj As ExpressionTree
8510      Set obj = Children(Index)
8511      Set GetChild = obj
End Function

' Setter method for children
Public Sub SetChild(Index As Long, obj As ExpressionTree)
8512      If NumChildren = 0 Then
8513          ReDim Preserve Children(Index)
8514          NumChildren = Index
8515      ElseIf Index > UBound(Children) Then
8516          ReDim Preserve Children(Index)
8517          NumChildren = Index
8518      End If
8519      Set Children(Index) = obj
End Sub

Public Property Get GetNumChildren() As Long
8520      GetNumChildren = NumChildren
End Property

' Clear all children from the node
Public Sub ClearChildren()
8521      NumChildren = 0
8522      ReDim Children(0)
End Sub

' Set the defaults for the node on creation
Private Sub Class_Initialize()
8523      NumChildren = 0
8524      LinearCoefficient = 0
8525      NodeLinearity = ExpressionTreeNonLinear
End Sub

' Test is the tree is empty
Public Function IsEmpty() As Boolean
8526      IsEmpty = NodeText = "0"
End Function


' Negate the tree
Public Function Negate() As ExpressionTree
8527      If NodeText <> "0" Then
              Dim NewTree As ExpressionTree
8528          Set NewTree = CreateTree("neg", ExpressionTreeOperator)
              
8529          NewTree.NodeLinearity = NodeLinearity
8530          NewTree.SetChild 1, Me
              
8531          Set Negate = NewTree
8532      Else
8533          Set Negate = Me
8534      End If
End Function

' Display the entire tree as a string for debugging
Public Function Display() As String
          Dim output As String
8535      If NodeType = ExpressionTreeOperator Then
8536          output = "("
8537      Else
8538          output = ""
8539      End If
          
          Dim i As Long
8540      For i = 1 To NumChildren
8541          output = output + Children(i).Display
8542      Next i
          
8543      output = output + NodeText
          
8544      If NodeType = ExpressionTreeOperator Then
8545          output = output + ") "
8546      Else
8547          output = output + " "
8548      End If
          
8549      Display = output
End Function

' Display the linearity of a node for debugging
Function DisplayLinearity() As String
8550      Select Case NodeLinearity
          Case ExpressionTreeNonLinear
8551          DisplayLinearity = "N"
8552      Case ExpressionTreeLinear
8553          DisplayLinearity = "L"
8554      Case ExpressionTreeConstant
8555          DisplayLinearity = "C"
8556      End Select
End Function

' Converts the tree to .nl expression tree format
Public Function ConvertToNL(Optional CommentIndent As Integer = 0) As String
          ' Construct comment
          Dim Comment As String, i As Long
8557      For i = 1 To CommentIndent
8558          Comment = Comment + " "
8559      Next i
8560      Comment = Comment + NodeText + " " + DisplayLinearity()
          
          ' Add current node to output
          Dim output As String
8561      AddNewLine output, FormatNL(NodeText, NodeType), Comment
          
          ' If n-ary operator, we need the number of operators on a newline
8562      If IsNAry(NodeText) Then
8564          Comment = ""
8565          For i = 1 To CommentIndent + 2
8566              Comment = Comment + " "
8567          Next i
8568          AddNewLine output, CStr(NumChildren), Comment & NumChildren & " arguments"
8570      End If
          
          ' Add children nodes to output
8572      For i = 1 To NumChildren
8573          output = output + Children(i).ConvertToNL(CommentIndent + 4)
8574      Next i
          
8576      ConvertToNL = output
End Function

' Extract all variables from the ExpressionTree into a LinearConstraintNL
Sub ExtractVariables(constraint As LinearConstraintNL)
8577      If NodeType = ExpressionTreeVariable Then
              Dim Index As Long
8578          Index = VariableIndex(NodeText)
8579          If Not constraint.VariablePresent(Index) Then
8580              constraint.VariablePresent(Index) = True
8581          End If
8582      End If
          
          Dim i As Long
8583      For i = 1 To NumChildren
8584          Children(i).ExtractVariables constraint
8585      Next i

End Sub

' Traverse the tree and mark all nodes as Constant, Linear or Non-linear
Sub MarkLinearity()
          ' First process all children
          Dim i As Long
8586      For i = 1 To NumChildren
8587          Children(i).MarkLinearity
8588      Next i

8589      Select Case NodeType
          ' A constant node is constant
          Case ExpressionTreeNumber, ExpressionTreeString
8590          NodeLinearity = ExpressionTreeConstant
          
          ' A variable is linear
8591      Case ExpressionTreeVariable
8592          NodeLinearity = ExpressionTreeLinear
              
8593      Case ExpressionTreeOperator
8594          Select Case NodeText
              Case "plus", "minus"
                  ' Non-linear if either child is non-linear
8595              If Children(1).NodeLinearity = ExpressionTreeNonLinear Or Children(2).NodeLinearity = ExpressionTreeNonLinear Then
8596                  NodeLinearity = ExpressionTreeNonLinear
                  ' Otherwise linear if either child is linear
8597              ElseIf Children(1).NodeLinearity = ExpressionTreeLinear Or Children(2).NodeLinearity = ExpressionTreeLinear Then
8598                  NodeLinearity = ExpressionTreeLinear
                  ' Otherwise constant
8599              Else
8600                  NodeLinearity = ExpressionTreeConstant
8601              End If
                  
8602          Case "mult"
                  ' Non-linear if either child is non-linear or both are linear
8603              If Children(1).NodeLinearity = ExpressionTreeNonLinear Or Children(2).NodeLinearity = ExpressionTreeNonLinear Or _
                     (Children(1).NodeLinearity = ExpressionTreeLinear And Children(2).NodeLinearity = ExpressionTreeLinear) Then
8604                  NodeLinearity = ExpressionTreeNonLinear
                  ' Otherwise linear if either child is linear
8605              ElseIf Children(1).NodeLinearity = ExpressionTreeLinear Or Children(2).NodeLinearity = ExpressionTreeLinear Then
8606                  NodeLinearity = ExpressionTreeLinear
                  ' Otherwise constant
8607              Else
8608                  NodeLinearity = ExpressionTreeConstant
8609              End If
                  
8610          Case "div"
                  ' If denominator is not constant, expression is non-linear
8611              If Not Children(2).NodeLinearity = ExpressionTreeConstant Then
8612                  NodeLinearity = ExpressionTreeNonLinear
                  ' Otherwise denominator is constant and expression is same linearity as numerator
8613              Else
8614                  NodeLinearity = Children(1).NodeLinearity
8615              End If
              
8616          Case "pow"
                  ' If exponent is not constant, then non-linear
8617              If Not Children(2).NodeLinearity = ExpressionTreeConstant Then
8618                 NodeLinearity = ExpressionTreeNonLinear
                  ' Otherwise non-linear if linear variable raised to power other than one
8619              ElseIf Children(1).NodeLinearity = ExpressionTreeLinear And Children(2).NodeText <> 1 Then
8620                 NodeLinearity = ExpressionTreeNonLinear
                  ' Otherwise expression is same linearity as base
8621              Else
8622                 NodeLinearity = Children(1).NodeLinearity
8623              End If
              
8624          Case "neg"
                  ' Same linearity as child
8625              NodeLinearity = Children(1).NodeLinearity
                  
              Case "if"
                  If Children(1).NodeLinearity = ExpressionTreeConstant Then
                      ' We know the outcome of the node already. Replace it with the result
                      Children(1).PullUpConstants
                      If Children(1).NodeText = 1 Then
                          ' Result was true. Replace with second node
                          SetTreeToChild 2
                      Else
                          ' Result was false. Replace with third node
                          SetTreeToChild 3
                      End If
                  Else
                      ' We can't tell which branch of the if will be taken
                      NodeLinearity = ExpressionTreeNonLinear
                  End If
                  
                  
8626          Case Else
                  ' Other functions
                  ' If all arguments const, then node is const
                  ' Otherwise non-linear
                  ' This is a catch-all, specific functions might need more detail
8627              For i = 1 To NumChildren
8628                  If Not Children(i).NodeLinearity = ExpressionTreeConstant Then
8629                      NodeLinearity = ExpressionTreeNonLinear
8630                      Exit Sub
8631                  End If
8632              Next i
8633              NodeLinearity = ExpressionTreeConstant
                  
8634          End Select
8635      End Select
End Sub

Public Sub PullUpConstants()
          ' First pull all constants in children
          Dim i As Long
8636      For i = 1 To NumChildren
8637          Children(i).PullUpConstants
8638      Next i

8639      If NodeLinearity = ExpressionTreeConstant And NodeType = ExpressionTreeOperator Then
8640          EvaluateConstantOperator
8641      End If
End Sub

' Evaluates a constant operator node and replaces node and all children with a numeric constant
Public Sub EvaluateConstantOperator()
          Dim i As Long, Values As New Collection, HasString As Boolean
          HasString = False
8642      For i = 1 To NumChildren
              If Children(i).NodeType = ExpressionTreeNumber Then
8643              Values.Add CDbl(Children(i).NodeText)
              Else
                  Values.Add Children(i).NodeText
                  HasString = True
              End If
8644      Next i
          
          Dim NewValue As Double
          
          If HasString Then
              Select Case NodeText
              Case "lt"
                  If Values(1) < Values(2) Then
                      NewValue = 1
                  Else
                      NewValue = 0
                  End If
              
              Case "le"
                  If Values(1) <= Values(2) Then
                      NewValue = 1
                  Else
                      NewValue = 0
                  End If
                  
              Case "eq"
                  If Values(1) = Values(2) Then
                      NewValue = 1
                  Else
                      NewValue = 0
                  End If
                  
              Case "ge"
                  If Values(1) >= Values(2) Then
                      NewValue = 1
                  Else
                      NewValue = 0
                  End If
              
              Case "gt"
                  If Values(1) > Values(2) Then
                      NewValue = 1
                  Else
                      NewValue = 0
                  End If
              
              Case "ne"
                  If Values(1) <> Values(2) Then
                      NewValue = 1
                  Else
                      NewValue = 0
                  End If
              
              Case Else
                  NewValue = 0
                  Err.Raise OpenSolver_BuildError, "Building expression tree", "String evaluation of " & NodeText & " not implemented yet in collation of linear terms" & vbCrLf & "Please let us know about this at opensolver.org so we can fix it."
              End Select
          Else
          
8645          Select Case NodeText
          
              Case "plus"
8646              NewValue = Values(1) + Values(2)
          
8647          Case "minus"
8648              NewValue = Values(1) - Values(2)
          
8649          Case "mult"
8650              NewValue = Values(1) * Values(2)
          
8651          Case "div"
8652              NewValue = Values(1) / Values(2)
              
8653          Case "rem"
8654              NewValue = Values(1) Mod Values(2)
              
8655          Case "pow"
8656              NewValue = Values(1) ^ Values(2)
              
8657          Case "less"
                  ' Returns max(arg1 - arg2, 0)
8658              NewValue = Values(1) - Values(2)
8659              If NewValue < 0 Then NewValue = 0
              
8660          Case "min"
8661              NewValue = 10000000000000# ' a big number
8662              For i = 1 To NumChildren
8663                  If Values(i) < NewValue Then NewValue = Values(i)
8664              Next i
              
8665          Case "max"
8666              NewValue = -10000000000000# ' a small number
8667              For i = 1 To NumChildren
8668                  If Values(i) > NewValue Then NewValue = Values(i)
8669              Next i
              
8670          Case "floor"
8671              NewValue = Int(Values(1))
              
8672          Case "ceil"
8673              NewValue = WorksheetFunction.Ceiling(Values(1), 1)
              
8674          Case "abs"
8675              NewValue = Abs(Values(1))
              
8676          Case "neg"
8677              NewValue = -Values(1)
              
8678          Case "or"
                  ' Returns 1 if any arg is non-zero
8679              If Not IsZero(Values(1)) Or Not IsZero(Values(2)) Then
8680                  NewValue = 1
8681              Else
8682                  NewValue = 0
8683              End If
              
8684          Case "and"
                  ' Returns 1 if both args are non-zero
8685              If Not IsZero(Values(1)) And Not IsZero(Values(2)) Then
8686                  NewValue = 1
8687              Else
8688                  NewValue = 0
8689              End If
              
8690          Case "lt"
8691              If Values(1) < Values(2) Then
8692                  NewValue = 1
8693              Else
8694                  NewValue = 0
8695              End If
              
8696          Case "le"
8697              If Values(1) <= Values(2) Then
8698                  NewValue = 1
8699              Else
8700                  NewValue = 0
8701              End If
              
8702          Case "eq"
                  If IsZero(Values(1) - Values(2)) Then
8704                  NewValue = 1
8705              Else
8706                  NewValue = 0
8707              End If
              
8708          Case "ge"
8709              If Values(1) >= Values(2) Then
8710                  NewValue = 1
8711              Else
8712                  NewValue = 0
8713              End If
              
8714          Case "gt"
8715              If Values(1) > Values(2) Then
8716                  NewValue = 1
8717              Else
8718                  NewValue = 0
8719              End If
              
8720          Case "ne"
8721              If Not IsZero(Values(1) - Values(2)) Then
8722                  NewValue = 1
8723              Else
8724                  NewValue = 0
8725              End If
              
8726          Case "if"
8727              If Not IsZero(Values(1)) Then
8728                  NewValue = Values(2)
8729              Else
8730                  NewValue = Values(3)
8731              End If
              
8732          Case "not"
                  ' Turn zero into 1, and everything else to 0
8733              If IsZero(Values(1)) Then
8734                  NewValue = 1
8735              Else
8736                  NewValue = 0
8737              End If
              
8738          Case "tanh"
8739              NewValue = WorksheetFunction.Tanh(Values(1))
              
8740          Case "tan"
8741              NewValue = Tan(Values(1))
              
8742          Case "sqrt"
8743              NewValue = Sqr(Values(1))
              
8744          Case "sinh"
8745              NewValue = WorksheetFunction.Sinh(Values(1))
              
8746          Case "sin"
8747              NewValue = Sin(Values(1))
              
8748          Case "log10"
                  ' VBA log is the natural logarithm - convert to base 10
8749              NewValue = Log(Values(1)) / Log(10)
              
8750          Case "log"
8751              NewValue = Log(Values(1))
              
8752          Case "exp"
8753              NewValue = Exp(Values(1))
              
8754          Case "cosh"
8755              NewValue = WorksheetFunction.Cosh(Values(1))
              
8756          Case "cos"
8757              NewValue = Cos(Values(1))
              
8758          Case "atanh"
8759              NewValue = WorksheetFunction.Atanh(Values(1))
              
8760          Case "atan2"
8761              NewValue = WorksheetFunction.Atan2(Values(1), Values(2))
              
8762          Case "atan"
8763              NewValue = Atn(Values(1))
              
8764          Case "asinh"
8765              NewValue = WorksheetFunction.Asinh(Values(1))
              
8766          Case "asin"
8767              NewValue = WorksheetFunction.Asin(Values(1))
              
8768          Case "acosh"
8769              NewValue = WorksheetFunction.Acosh(Values(1))
              
8770          Case "acos"
8771              NewValue = WorksheetFunction.Acos(Values(1))
                  
8772          Case "sum"
8773              NewValue = 0
8774              For i = 1 To NumChildren
8775                  NewValue = NewValue + Values(i)
8776              Next i
                  
8777          Case "intdiv"
8778              NewValue = Values(1) \ Values(2)
                  
8779          Case "precision"
                  ' Rounds number to specified number of sig figs
                  ' Formula from http://stackoverflow.com/a/15976194
8780              NewValue = Round(Values(1), Values(2) - (Int(Log(Values(1)) / Log(10#)) + 1))
                  
8781          Case "round"
                  ' Rounds number to specified number of decimals
8782              NewValue = Round(Values(1), Values(2))
                  
8783          Case "trunc"
                  ' Truncates number to specified number of decimals
8784              NewValue = WorksheetFunction.RoundDown(Values(1), Values(2))
                  
8785          Case "count"
                  ' Counts number of non-zero arguments
8786              NewValue = 0
8787              For i = 1 To NumChildren
8788                  If Not IsZero(Values(i)) Then NewValue = NewValue + 1
8789              Next i
                  
8790          Case "numberof"
                  ' Counts number of numeric args that match the first
                  Dim NumToMatch As Double
8791              NumToMatch = Values(1)
8792              NewValue = 0
8793              For i = 2 To NumChildren
8794                  If IsZero(NumToMatch - Values(i)) Then NewValue = NewValue + 1
8795              Next i
                  
8796          Case "numberofs"
                  ' Counts number of string args that match the first
                  Dim StringToMatch As String
8797              StringToMatch = Values(1)
8798              NewValue = 0
8799              For i = 2 To NumChildren
8800                  If StringToMatch = Values(i) Then NewValue = NewValue + 1
8801              Next i
                  
8802          Case "ifs"
8803              If Not IsZero(Values(1)) Then
8804                  NewValue = Values(2)
8805              Else
8806                  NewValue = Values(3)
8807              End If
                  
8808          Case "and_n"
                  ' Returns 1 if all args are non-zero
8809              NewValue = 1
8810              For i = 1 To NumChildren
8811                  If IsZero(Values(i)) Then
8812                      NewValue = 0
8813                      GoTo FinishNode
8814                  End If
8815              Next i
                  
8816          Case "or_n"
                  ' Returns 1 if any args are non-zero
8817              NewValue = 0
8818              For i = 1 To NumChildren
8819                  If Not IsZero(Values(i)) Then
8820                      NewValue = 1
8821                      GoTo FinishNode
8822                  End If
8823              Next i
                  
8824          Case "implies"
8825              If Not IsZero(Values(1)) Then
8826                  NewValue = Values(2)
8827              Else
8828                  NewValue = Values(3)
8829              End If
                  
8830          Case "iff"
                  ' Returns 1 if both args are non-zero or both are zero (XNOR)
8831              If (Not IsZero(Values(1)) And Not IsZero(Values(2))) Or (IsZero(Values(1)) And IsZero(Values(2))) Then
8832                  NewValue = 1
8833              Else
8834                  NewValue = 0
8835              End If
                  
8836          Case "alldiff"
                  ' Returns 1 if all arguments are different
                  Dim Args As New Collection, Key As String
8837              For i = 1 To NumChildren
                      ' Create key using the var type so we distinguish between 1.23 and "1.23"
8838                  Key = CStr(VarType(Values(i))) & CStr(Values(i))
8839                  If TestKeyExists(Args, Key) Then
8840                      NewValue = 0 ' Duplicate found
8841                      GoTo FinishNode
8842                  Else
8843                      Args.Add 0, Key ' Value doesn't matter, we just want to track the key
8844                  End If
8845              Next i
                  
8846          Case Else
8847              NewValue = 0
8848              Err.Raise OpenSolver_BuildError, "Building expression tree", "Constant evaluation of " & NodeText & " not implemented yet in collation of linear terms" & vbCrLf & "Please let us know about this at opensolver.org so we can fix it."
8849          End Select
          End If
              
FinishNode:
          ' Replace subtree with a constant node
8850      NodeType = ExpressionTreeNumber
8851      NodeText = NewValue
8852      ClearChildren
End Sub

' Traverses the tree and prunes all linear or constant parts from the tree
Public Function PruneLinearTrees(LinearTrees As Collection, Optional NegativeTree As Boolean = False) As Boolean
          Dim CanPrune As Boolean, i As Long, NewTree As ExpressionTree
8853      CanPrune = False
          ' If the root node is linear or constant we can prune
8854      If NodeLinearity <> ExpressionTreeNonLinear Then
8855         CanPrune = True
          ' Otherwise if the root is a "plus" or "minus", it might have a linear part as one of the children
8856      ElseIf NodeText = "plus" Or NodeText = "minus" Then
             'Prune children individually and set to n0 nodes
              Dim PrunedFirst As Boolean, PrunedSecond As Boolean
              
8857          PrunedFirst = Children(1).PruneLinearTrees(LinearTrees)
8858          If NodeText = "minus" Then
8859              PrunedSecond = Children(2).PruneLinearTrees(LinearTrees, True)
8860          Else
8861              PrunedSecond = Children(2).PruneLinearTrees(LinearTrees)
8862          End If
              
              ' If first tree pruned, then set this node to the second child
8863          If PrunedFirst Then
                  ' Negate the tree if it was a minus node before setting
8864              If NodeText = "minus" Then
8865                  SetChild 2, Children(2).Negate
8866              End If
8867              SetTreeToChild 2
                  
              ' If second node pruned, then set this node to the first child
8868          ElseIf PrunedSecond Then
8869              SetTreeToChild 1
8870          End If
8871      End If
          
8872      If CanPrune Then
              ' Copy tree
8873          Set NewTree = Me.Clone
              
              ' If tree is being pruned from a "minus", we need to negate it
8874          If NegativeTree Then
8875              Set NewTree = NewTree.Negate
8876          End If

              ' Add tree to linear collection
8877          LinearTrees.Add NewTree
              
              ' Reset current tree to n0
8878          NodeText = 0
8879          NodeType = ExpressionTreeNumber
8880          NodeLinearity = ExpressionTreeConstant
8881          ClearChildren
              
8882          PruneLinearTrees = True
8883      Else
8884          PruneLinearTrees = False
8885      End If
              
End Function

' Sets an ExpressionTree to be one of its children
Sub SetTreeToChild(Index As Long)
          ' We can't clear all children of the current tree while referring to one of it's children
          ' We therefore need to make a copy of the child tree before we clear the tree and set it to the child
          Dim NewTree As ExpressionTree
8886      Set NewTree = Children(Index).Clone
8887      SetExpressionTree NewTree
End Sub

' Creates a copy of an ExpressionTree
Public Function Clone() As ExpressionTree
          Dim NewTree As ExpressionTree
8888      Set NewTree = CreateTree(NodeText, NodeType)
8889      NewTree.NodeLinearity = NodeLinearity
8890      NewTree.LinearCoefficient = LinearCoefficient
8891      NewTree.LinearValue = LinearValue
          
          Dim i As Long
8892      For i = 1 To NumChildren
8893          NewTree.SetChild i, Children(i)
8894      Next i
          
8895      Set Clone = NewTree
End Function

' Sets the current ExpressionTree to be the same as another
Sub SetExpressionTree(NewTree As ExpressionTree)
8896      NodeText = NewTree.NodeText
8897      NodeType = NewTree.NodeType
8898      NodeLinearity = NewTree.NodeLinearity
8899      LinearCoefficient = NewTree.LinearCoefficient
8900      LinearValue = NewTree.LinearValue
          
8901      ClearChildren
          
          Dim i As Long
8902      For i = 1 To NewTree.GetNumChildren
8903          SetChild i, NewTree.GetChild(i)
8904      Next i
End Sub

' Converts a linear tree into a LinearConstraintNL and Double
Public Sub ConvertLinearTreeToConstraint(constraint As LinearConstraintNL, constant As Double)
          ' Remove everything but "plus", variables, and constants
8905      CollateLinearTerms
          
          ' Add linear terms into constraint/constant
8906      AddLinearTreeToConstraint constraint, constant
End Sub

' Traverses a linear tree and replaces all operators other than "plus"
' We want a tree with constants, variables, and "plus" nodes only
Sub CollateLinearTerms()

          ' First collate all children
          Dim i As Long
8907      For i = 1 To NumChildren
8908          Children(i).CollateLinearTerms
8909      Next i

8910      Select Case NodeLinearity
          ' If the node if constant, we want to evaluate it
          Case ExpressionTreeConstant
8911          Select Case NodeType
              ' Constant expression on a number node - we read the number
              Case ExpressionTreeNumber
8912              LinearValue = CDbl(NodeText)
              
              ' Constant expression under operator - we can evaluate it and discard all children
8913          Case ExpressionTreeOperator
8914              EvaluateConstantOperator
8915              LinearValue = CDbl(NodeText)
8916          End Select
              
          ' A linear node - we need to update the coefficient on the linear term
8917      Case ExpressionTreeLinear
8918          Select Case NodeType
              ' Variable - Intitialise coefficient at 1 if hasn't been already
              Case ExpressionTreeVariable
8919              If LinearCoefficient = 0 Then
8920                  LinearCoefficient = 1
8921              End If
              
              ' Linear operator node - we need to collate the constant and variable terms into a single variable node if possible
8922          Case ExpressionTreeOperator
8923              Select Case NodeText
                  
                  ' Do nothing on a plus node. This is just to escape the Else case
                  Case "plus"
                  
                  ' If we have a minus node, convert it to a plus and flip the sign on the second operator
8924              Case "minus"
8925                  Children(2).CollateLinearNegative
8926                  NodeText = "plus"
                      
                  ' Negate the coefficient on the variable node and remove the "neg" node
8927              Case "neg"
8928                  Children(1).MultiplyByConstant -1
8929                  SetTreeToChild 1
                          
                  ' There is a constant node and a linear node attached to this mult.
                  ' We multiply the constant through the linear node (which may be a "plus" tree)
8930              Case "mult"
8931                  If Children(1).NodeType = ExpressionTreeConstant Then
8932                      Children(2).MultiplyByConstant Children(1).LinearValue
8933                      SetTreeToChild 2
8934                  Else ' Child 2 is the constant node
8935                      Children(1).MultiplyByConstant Children(2).LinearValue
8936                      SetTreeToChild 1
8937                  End If
                  
                  ' First node is variable
8938              Case "div"
8939                  Children(1).MultiplyByConstant 1 / Children(2).LinearValue
8940                  SetTreeToChild 1
                      
                  ' Second node must be 1 (as we know the expression is linear)
8941              Case "pow"
8942                  SetTreeToChild 1
          
8943              Case Else
8944                  Err.Raise OpenSolver_BuildError, "Collating linear terms", "Unknown linear operator " & NodeText & vbCrLf & "Please let us know about this at opensolver.org so we can fix it."
8945              End Select
8946          End Select
8947      End Select
          
8948      Exit Sub
          
End Sub

' Multiplies all nodes in the linear tree by a constant
Public Sub MultiplyByConstant(constant As Double)
8949      Select Case NodeType
          Case ExpressionTreeConstant
8950          LinearValue = LinearValue * constant
8951          NodeText = LinearValue
          
8952      Case ExpressionTreeVariable
8953          LinearCoefficient = LinearCoefficient * constant
              
8954      Case ExpressionTreeOperator ' plus
8955          Children(1).MultiplyByConstant constant
8956          Children(2).MultiplyByConstant constant
              
8957      End Select
End Sub

' Flips the sign on variable or constant
Public Sub CollateLinearNegative()
8958      If NodeType = ExpressionTreeVariable Then
8959          LinearCoefficient = -LinearCoefficient
8960      Else ' NodeType = ExpressionTreeNumber
8961          If NodeText <> "0" Then
8962              LinearValue = -CDbl(NodeText)
8963              NodeText = LinearValue
8964          End If
8965      End If
End Sub

' Convert a collated linear tree into a LinearConstraintNL and Double constant
Sub AddLinearTreeToConstraint(constraint As LinearConstraintNL, constant As Double)
          ' By this point, only "plus", variables and constants remain in the tree.
          ' We traverse the tree and add each variable we find to the constraint,
          ' and each constant value to the constant total
          
8966      Select Case NodeType
          ' Add variable coefficient to constraint
          Case ExpressionTreeVariable
              Dim VarIndex As Long
8967          VarIndex = VariableIndex(NodeText)
8968          constraint.Coefficient(VarIndex) = constraint.Coefficient(VarIndex) + LinearCoefficient
          
          ' Add constant
8969      Case ExpressionTreeNumber
8970          constant = constant + LinearValue
          
          ' Recurse on all "plus" nodes
8971      Case ExpressionTreeOperator
8972          If NodeText = "plus" Then
                  Dim i As Long
8973              For i = 1 To NumChildren
8974                  Children(i).AddLinearTreeToConstraint constraint, constant
8975              Next i
8976          Else
8977              Err.Raise OpenSolver_BuildError, "Collating linear terms", "Operator that isn't plus at the last stage of the linear tree pruning: " & NodeText & vbCrLf & "Please let us know about this at opensolver.org so we can fix it."
8978          End If
8979      End Select
End Sub
