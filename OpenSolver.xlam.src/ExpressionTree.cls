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
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim obj As ExpressionTree
3         Set obj = Children(Index)
4         Set GetChild = obj

ExitFunction:
5         If RaiseError Then RethrowError
6         Exit Function

ErrorHandler:
7         If Not ReportError("ExpressionTree", "GetChild") Then Resume
8         RaiseError = True
9         GoTo ExitFunction
End Function

' Setter method for children
Public Sub SetChild(Index As Long, obj As ExpressionTree)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If NumChildren = 0 Then
4             ReDim Preserve Children(Index)
5             NumChildren = Index
6         ElseIf Index > UBound(Children) Then
7             ReDim Preserve Children(Index)
8             NumChildren = Index
9         End If
10        Set Children(Index) = obj

ExitSub:
11        If RaiseError Then RethrowError
12        Exit Sub

ErrorHandler:
13        If Not ReportError("ExpressionTree", "SetChild") Then Resume
14        RaiseError = True
15        GoTo ExitSub
End Sub

Public Property Get GetNumChildren() As Long
1         GetNumChildren = NumChildren
End Property

' Clear all children from the node
Public Sub ClearChildren()
1         NumChildren = 0
2         ReDim Children(0)
End Sub

' Set the defaults for the node on creation
Private Sub Class_Initialize()
1         NumChildren = 0
2         LinearCoefficient = 0
3         NodeLinearity = ExpressionTreeNonLinear
End Sub

' Test is the tree is empty
Public Function IsEmpty() As Boolean
1         IsEmpty = NodeText = "0"
End Function


' Negate the tree
Public Function Negate() As ExpressionTree
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If NodeText <> "0" Then
              Dim NewTree As ExpressionTree
4             Set NewTree = CreateTree("neg", ExpressionTreeOperator)
              
5             NewTree.NodeLinearity = NodeLinearity
6             NewTree.SetChild 1, Me
              
7             Set Negate = NewTree
8         Else
9             Set Negate = Me
10        End If

ExitFunction:
11        If RaiseError Then RethrowError
12        Exit Function

ErrorHandler:
13        If Not ReportError("ExpressionTree", "Negate") Then Resume
14        RaiseError = True
15        GoTo ExitFunction
End Function

' Display the entire tree as a string for debugging
Public Function Display() As String
          Dim output As String
1         If NodeType = ExpressionTreeOperator Then
2             output = "("
3         Else
4             output = vbNullString
5         End If
          
          Dim i As Long
6         For i = 1 To NumChildren
7             output = output + Children(i).Display
8         Next i
          
9         output = output + NodeText
          
10        If NodeType = ExpressionTreeOperator Then
11            output = output + ") "
12        Else
13            output = output + " "
14        End If
          
15        Display = output
End Function

' Display the linearity of a node for debugging
Function DisplayLinearity() As String
1         Select Case NodeLinearity
          Case ExpressionTreeNonLinear: DisplayLinearity = "N"
2         Case ExpressionTreeLinear:    DisplayLinearity = "L"
3         Case ExpressionTreeConstant:  DisplayLinearity = "C"
4         End Select
End Function

' Converts the tree to .nl expression tree format
Public Sub ConvertToNL(FileNum As Long, Optional CommentIndent As Integer = 0)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          ' Add current node to output
3         OutputLine FileNum, _
              FormatNL(NodeText, NodeType), _
              Space(CommentIndent) & NodeText & " " & DisplayLinearity()
          
          ' If n-ary operator, we need the number of operators on a newline
4         If IsNAry(NodeText) Then
5             OutputLine FileNum, _
                  CStr(NumChildren), _
                  Space(CommentIndent + 2) & NumChildren & " arguments"
6         End If
          
          ' Add children nodes to output
          Dim i As Long
7         For i = 1 To NumChildren
8             Children(i).ConvertToNL FileNum, CommentIndent + 4
9         Next i

ExitFunction:
10        If RaiseError Then RethrowError
11        Exit Sub

ErrorHandler:
12        If Not ReportError("ExpressionTree", "ConvertToNL") Then Resume
13        RaiseError = True
14        GoTo ExitFunction
End Sub

' Extract all variables from the ExpressionTree into a Dictionary
Sub ExtractVariables(constraint As Dictionary)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If NodeType = ExpressionTreeVariable Then
              Dim Index As Long
4             Index = VariableIndex.Item(NodeText)
5             If Not constraint.Exists(Index) Then
6                 constraint.Add Key:=Index, Item:=0
7             End If
8         End If
          
          Dim i As Long
9         For i = 1 To NumChildren
10            Children(i).ExtractVariables constraint
11        Next i

ExitSub:
12        If RaiseError Then RethrowError
13        Exit Sub

ErrorHandler:
14        If Not ReportError("ExpressionTree", "ExtractVariables") Then Resume
15        RaiseError = True
16        GoTo ExitSub
End Sub

' Traverse the tree and mark all nodes as Constant, Linear or Non-linear
Sub MarkLinearity()
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          ' First process all children
          Dim i As Long
3         For i = 1 To NumChildren
4             Children(i).MarkLinearity
5         Next i

6         Select Case NodeType
          ' A constant node is constant
          Case ExpressionTreeNumber, ExpressionTreeString
7             NodeLinearity = ExpressionTreeConstant
          
          ' A variable is linear
8         Case ExpressionTreeVariable
9             NodeLinearity = ExpressionTreeLinear
              
10        Case ExpressionTreeOperator
11            Select Case NodeText
              Case "plus", "minus"
                  ' Non-linear if either child is non-linear
12                If Children(1).NodeLinearity = ExpressionTreeNonLinear Or Children(2).NodeLinearity = ExpressionTreeNonLinear Then
13                    NodeLinearity = ExpressionTreeNonLinear
                  ' Otherwise linear if either child is linear
14                ElseIf Children(1).NodeLinearity = ExpressionTreeLinear Or Children(2).NodeLinearity = ExpressionTreeLinear Then
15                    NodeLinearity = ExpressionTreeLinear
                  ' Otherwise constant
16                Else
17                    NodeLinearity = ExpressionTreeConstant
18                End If
                  
19            Case "mult"
                  ' Non-linear if either child is non-linear or both are linear
20                If Children(1).NodeLinearity = ExpressionTreeNonLinear Or Children(2).NodeLinearity = ExpressionTreeNonLinear Or _
                     (Children(1).NodeLinearity = ExpressionTreeLinear And Children(2).NodeLinearity = ExpressionTreeLinear) Then
21                    NodeLinearity = ExpressionTreeNonLinear
                  ' Otherwise linear if either child is linear
22                ElseIf Children(1).NodeLinearity = ExpressionTreeLinear Or Children(2).NodeLinearity = ExpressionTreeLinear Then
23                    NodeLinearity = ExpressionTreeLinear
                  ' Otherwise constant
24                Else
25                    NodeLinearity = ExpressionTreeConstant
26                End If
                  
27            Case "div"
                  ' If denominator is not constant, expression is non-linear
28                If Not Children(2).NodeLinearity = ExpressionTreeConstant Then
29                    NodeLinearity = ExpressionTreeNonLinear
                  ' Otherwise denominator is constant and expression is same linearity as numerator
30                Else
31                    NodeLinearity = Children(1).NodeLinearity
32                End If
              
33            Case "pow"
                  ' If exponent is not constant, then non-linear
34                If Not Children(2).NodeLinearity = ExpressionTreeConstant Then
35                   NodeLinearity = ExpressionTreeNonLinear
                  ' Otherwise non-linear if linear variable raised to power other than one
36                ElseIf Children(1).NodeLinearity = ExpressionTreeLinear And Children(2).NodeText <> "1" Then
37                   NodeLinearity = ExpressionTreeNonLinear
                  ' Otherwise expression is same linearity as base
38                Else
39                   NodeLinearity = Children(1).NodeLinearity
40                End If
              
41            Case "neg"
                  ' Same linearity as child
42                NodeLinearity = Children(1).NodeLinearity
                  
43            Case "if"
44                If Children(1).NodeLinearity = ExpressionTreeConstant Then
                      ' We know the outcome of the node already. Replace it with the result
45                    Children(1).PullUpConstants
46                    If Children(1).NodeText = 1 Then
                          ' Result was true. Replace with second node
47                        SetTreeToChild 2
48                    Else
                          ' Result was false. Replace with third node
49                        SetTreeToChild 3
50                    End If
51                Else
                      ' We can't tell which branch of the if will be taken
52                    NodeLinearity = ExpressionTreeNonLinear
53                End If
                  
                  
54            Case Else
                  ' Other functions
                  ' If all arguments const, then node is const
                  ' Otherwise non-linear
                  ' This is a catch-all, specific functions might need more detail
55                For i = 1 To NumChildren
56                    If Not Children(i).NodeLinearity = ExpressionTreeConstant Then
57                        NodeLinearity = ExpressionTreeNonLinear
58                        GoTo ExitSub
59                    End If
60                Next i
61                NodeLinearity = ExpressionTreeConstant
                  
62            End Select
63        End Select

ExitSub:
64        If RaiseError Then RethrowError
65        Exit Sub

ErrorHandler:
66        If Not ReportError("ExpressionTree", "MarkLinearity") Then Resume
67        RaiseError = True
68        GoTo ExitSub
End Sub

Public Sub PullUpConstants()
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          ' First pull all constants in children
          Dim i As Long
3         For i = 1 To NumChildren
4             Children(i).PullUpConstants
5         Next i

6         If NodeLinearity = ExpressionTreeConstant And NodeType = ExpressionTreeOperator Then
7             EvaluateConstantOperator
8         End If

ExitSub:
9         If RaiseError Then RethrowError
10        Exit Sub

ErrorHandler:
11        If Not ReportError("ExpressionTree", "PullUpConstants") Then Resume
12        RaiseError = True
13        GoTo ExitSub
End Sub

' Evaluates a constant operator node and replaces node and all children with a numeric constant
Public Sub EvaluateConstantOperator()
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim i As Long, Values As New Collection, HasString As Boolean
3         HasString = False
4         For i = 1 To NumChildren
5             If Children(i).NodeType = ExpressionTreeNumber Then
6                 Values.Add Val(Children(i).NodeText)
7             Else
8                 Values.Add Children(i).NodeText
9                 HasString = True
10            End If
11        Next i
          
          Dim NewValue As Double
          
12        If HasString Then
13            Select Case NodeText
              Case "lt"
14                If Values(1) < Values(2) Then
15                    NewValue = 1
16                Else
17                    NewValue = 0
18                End If
              
19            Case "le"
20                If Values(1) <= Values(2) Then
21                    NewValue = 1
22                Else
23                    NewValue = 0
24                End If
                  
25            Case "eq"
26                If Values(1) = Values(2) Then
27                    NewValue = 1
28                Else
29                    NewValue = 0
30                End If
                  
31            Case "ge"
32                If Values(1) >= Values(2) Then
33                    NewValue = 1
34                Else
35                    NewValue = 0
36                End If
              
37            Case "gt"
38                If Values(1) > Values(2) Then
39                    NewValue = 1
40                Else
41                    NewValue = 0
42                End If
              
43            Case "ne"
44                If Values(1) <> Values(2) Then
45                    NewValue = 1
46                Else
47                    NewValue = 0
48                End If
              
49            Case Else
50                NewValue = 0
51                RaiseGeneralError "String evaluation of " & NodeText & " not implemented yet in collation of linear terms" & vbCrLf & "Please let us know about this so we can fix it."
52            End Select
53        Else
          
54            Select Case NodeText
          
              Case "plus"
55                NewValue = Values(1) + Values(2)
          
56            Case "minus"
57                NewValue = Values(1) - Values(2)
          
58            Case "mult"
59                NewValue = Values(1) * Values(2)
          
60            Case "div"
61                NewValue = Values(1) / Values(2)
              
62            Case "rem"
63                NewValue = Values(1) Mod Values(2)
              
64            Case "pow"
65                NewValue = Values(1) ^ Values(2)
              
66            Case "less"
                  ' Returns max(arg1 - arg2, 0)
67                NewValue = Values(1) - Values(2)
68                If NewValue < 0 Then NewValue = 0
              
69            Case "min"
70                NewValue = 10000000000000# ' a big number
71                For i = 1 To NumChildren
72                    If Values(i) < NewValue Then NewValue = Values(i)
73                Next i
              
74            Case "max"
75                NewValue = -10000000000000# ' a small number
76                For i = 1 To NumChildren
77                    If Values(i) > NewValue Then NewValue = Values(i)
78                Next i
              
79            Case "floor"
80                NewValue = Int(Values(1))
              
81            Case "ceil"
82                NewValue = WorksheetFunction.Ceiling(Values(1), 1)
              
83            Case "abs"
84                NewValue = Abs(Values(1))
              
85            Case "neg"
86                NewValue = -Values(1)
              
87            Case "or"
                  ' Returns 1 if any arg is non-zero
88                If Not IsZero(Values(1)) Or Not IsZero(Values(2)) Then
89                    NewValue = 1
90                Else
91                    NewValue = 0
92                End If
              
93            Case "and"
                  ' Returns 1 if both args are non-zero
94                If Not IsZero(Values(1)) And Not IsZero(Values(2)) Then
95                    NewValue = 1
96                Else
97                    NewValue = 0
98                End If
              
99            Case "lt"
100               If Values(1) < Values(2) Then
101                   NewValue = 1
102               Else
103                   NewValue = 0
104               End If
              
105           Case "le"
106               If Values(1) <= Values(2) Then
107                   NewValue = 1
108               Else
109                   NewValue = 0
110               End If
              
111           Case "eq"
112               If IsZero(Values(1) - Values(2)) Then
113                   NewValue = 1
114               Else
115                   NewValue = 0
116               End If
              
117           Case "ge"
118               If Values(1) >= Values(2) Then
119                   NewValue = 1
120               Else
121                   NewValue = 0
122               End If
              
123           Case "gt"
124               If Values(1) > Values(2) Then
125                   NewValue = 1
126               Else
127                   NewValue = 0
128               End If
              
129           Case "ne"
130               If Not IsZero(Values(1) - Values(2)) Then
131                   NewValue = 1
132               Else
133                   NewValue = 0
134               End If
              
135           Case "if"
136               If Not IsZero(Values(1)) Then
137                   NewValue = Values(2)
138               Else
139                   NewValue = Values(3)
140               End If
              
141           Case "not"
                  ' Turn zero into 1, and everything else to 0
142               If IsZero(Values(1)) Then
143                   NewValue = 1
144               Else
145                   NewValue = 0
146               End If
              
147           Case "tanh"
148               NewValue = WorksheetFunction.Tanh(Values(1))
              
149           Case "tan"
150               NewValue = Tan(Values(1))
              
151           Case "sqrt"
152               NewValue = Sqr(Values(1))
              
153           Case "sinh"
154               NewValue = WorksheetFunction.Sinh(Values(1))
              
155           Case "sin"
156               NewValue = Sin(Values(1))
              
157           Case "log10"
                  ' VBA log is the natural logarithm - convert to base 10
158               NewValue = Log(Values(1)) / Log(10)
              
159           Case "log"
160               NewValue = Log(Values(1))
              
161           Case "exp"
162               NewValue = Exp(Values(1))
              
163           Case "cosh"
164               NewValue = WorksheetFunction.Cosh(Values(1))
              
165           Case "cos"
166               NewValue = Cos(Values(1))
              
167           Case "atanh"
168               NewValue = WorksheetFunction.Atanh(Values(1))
              
169           Case "atan2"
170               NewValue = WorksheetFunction.Atan2(Values(1), Values(2))
              
171           Case "atan"
172               NewValue = Atn(Values(1))
              
173           Case "asinh"
174               NewValue = WorksheetFunction.Asinh(Values(1))
              
175           Case "asin"
176               NewValue = WorksheetFunction.Asin(Values(1))
              
177           Case "acosh"
178               NewValue = WorksheetFunction.Acosh(Values(1))
              
179           Case "acos"
180               NewValue = WorksheetFunction.Acos(Values(1))
                  
181           Case "sum"
182               NewValue = 0
183               For i = 1 To NumChildren
184                   NewValue = NewValue + Values(i)
185               Next i
                  
186           Case "intdiv"
187               NewValue = Values(1) \ Values(2)
                  
188           Case "precision"
                  ' Rounds number to specified number of sig figs
                  ' Formula from http://stackoverflow.com/a/15976194
189               NewValue = Round(Values(1), Values(2) - (Int(Log(Values(1)) / Log(10#)) + 1))
                  
190           Case "round"
                  ' Rounds number to specified number of decimals
191               NewValue = Round(Values(1), Values(2))
                  
192           Case "trunc"
                  ' Truncates number to specified number of decimals
193               NewValue = WorksheetFunction.RoundDown(Values(1), Values(2))
                  
194           Case "count"
                  ' Counts number of non-zero arguments
195               NewValue = 0
196               For i = 1 To NumChildren
197                   If Not IsZero(Values(i)) Then NewValue = NewValue + 1
198               Next i
                  
199           Case "numberof"
                  ' Counts number of numeric args that match the first
                  Dim NumToMatch As Double
200               NumToMatch = Values(1)
201               NewValue = 0
202               For i = 2 To NumChildren
203                   If IsZero(NumToMatch - Values(i)) Then NewValue = NewValue + 1
204               Next i
                  
205           Case "numberofs"
                  ' Counts number of string args that match the first
                  Dim StringToMatch As String
206               StringToMatch = Values(1)
207               NewValue = 0
208               For i = 2 To NumChildren
209                   If StringToMatch = Values(i) Then NewValue = NewValue + 1
210               Next i
                  
211           Case "ifs"
212               If Not IsZero(Values(1)) Then
213                   NewValue = Values(2)
214               Else
215                   NewValue = Values(3)
216               End If
                  
217           Case "and_n"
                  ' Returns 1 if all args are non-zero
218               NewValue = 1
219               For i = 1 To NumChildren
220                   If IsZero(Values(i)) Then
221                       NewValue = 0
222                       GoTo FinishNode
223                   End If
224               Next i
                  
225           Case "or_n"
                  ' Returns 1 if any args are non-zero
226               NewValue = 0
227               For i = 1 To NumChildren
228                   If Not IsZero(Values(i)) Then
229                       NewValue = 1
230                       GoTo FinishNode
231                   End If
232               Next i
                  
233           Case "implies"
234               If Not IsZero(Values(1)) Then
235                   NewValue = Values(2)
236               Else
237                   NewValue = Values(3)
238               End If
                  
239           Case "iff"
                  ' Returns 1 if both args are non-zero or both are zero (XNOR)
240               If (Not IsZero(Values(1)) And Not IsZero(Values(2))) Or (IsZero(Values(1)) And IsZero(Values(2))) Then
241                   NewValue = 1
242               Else
243                   NewValue = 0
244               End If
                  
245           Case "alldiff"
                  ' Returns 1 if all arguments are different
                  Dim Args As New Collection, Key As String
246               For i = 1 To NumChildren
                      ' Create key using the var type so we distinguish between 1.23 and "1.23"
247                   Key = CStr(VarType(Values(i))) & CStr(Values(i))
248                   If TestKeyExists(Args, Key) Then
249                       NewValue = 0 ' Duplicate found
250                       GoTo FinishNode
251                   Else
252                       Args.Add 0, Key ' Value doesn't matter, we just want to track the key
253                   End If
254               Next i
                  
255           Case Else
256               NewValue = 0
257               RaiseGeneralError "Constant evaluation of " & NodeText & " not implemented yet in collation of linear terms" & vbCrLf & "Please let us know about this so we can fix it."
258           End Select
259       End If
              
FinishNode:
          ' Replace subtree with a constant node
260       NodeType = ExpressionTreeNumber
261       NodeText = StrExNoPlus(NewValue)
262       ClearChildren

ExitSub:
263       If RaiseError Then RethrowError
264       Exit Sub

ErrorHandler:
265       If Not ReportError("ExpressionTree", "EvaluateConstantOperator") Then Resume
266       RaiseError = True
267       GoTo ExitSub
End Sub

' Traverses the tree and prunes all linear or constant parts from the tree
Public Function PruneLinearTrees(LinearTrees As Collection, Optional NegativeTree As Boolean = False) As Boolean
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim CanPrune As Boolean, NewTree As ExpressionTree
3         CanPrune = False
          ' If the root node is linear or constant we can prune
4         If NodeLinearity <> ExpressionTreeNonLinear Then
5            CanPrune = True
          ' Otherwise if the root is a "plus" or "minus", it might have a linear part as one of the children
6         ElseIf NodeText = "plus" Or NodeText = "minus" Then
             'Prune children individually and set to n0 nodes
              Dim PrunedFirst As Boolean, PrunedSecond As Boolean
              
7             PrunedFirst = Children(1).PruneLinearTrees(LinearTrees)
8             If NodeText = "minus" Then
9                 PrunedSecond = Children(2).PruneLinearTrees(LinearTrees, True)
10            Else
11                PrunedSecond = Children(2).PruneLinearTrees(LinearTrees)
12            End If
              
              ' If first tree pruned, then set this node to the second child
13            If PrunedFirst Then
                  ' Negate the tree if it was a minus node before setting
14                If NodeText = "minus" Then
15                    SetChild 2, Children(2).Negate
16                End If
17                SetTreeToChild 2
                  
              ' If second node pruned, then set this node to the first child
18            ElseIf PrunedSecond Then
19                SetTreeToChild 1
20            End If
21        End If
          
22        If CanPrune Then
              ' Copy tree
23            Set NewTree = Me.Clone
              
              ' If tree is being pruned from a "minus", we need to negate it
24            If NegativeTree Then
25                Set NewTree = NewTree.Negate
26            End If

              ' Add tree to linear collection
27            LinearTrees.Add NewTree
              
              ' Reset current tree to n0
28            NodeText = 0
29            NodeType = ExpressionTreeNumber
30            NodeLinearity = ExpressionTreeConstant
31            ClearChildren
              
32            PruneLinearTrees = True
33        Else
34            PruneLinearTrees = False
35        End If

ExitFunction:
36        If RaiseError Then RethrowError
37        Exit Function

ErrorHandler:
38        If Not ReportError("ExpressionTree", "PruneLinearTrees") Then Resume
39        RaiseError = True
40        GoTo ExitFunction
End Function

' Sets an ExpressionTree to be one of its children
Sub SetTreeToChild(Index As Long)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          ' We can't clear all children of the current tree while referring to one of it's children
          ' We therefore need to make a copy of the child tree before we clear the tree and set it to the child
          Dim NewTree As ExpressionTree
3         Set NewTree = Children(Index).Clone
4         SetExpressionTree NewTree

ExitSub:
5         If RaiseError Then RethrowError
6         Exit Sub

ErrorHandler:
7         If Not ReportError("ExpressionTree", "SetTreeToChild") Then Resume
8         RaiseError = True
9         GoTo ExitSub
End Sub

' Creates a copy of an ExpressionTree
Public Function Clone() As ExpressionTree
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim NewTree As ExpressionTree
3         Set NewTree = CreateTree(NodeText, NodeType)
4         NewTree.NodeLinearity = NodeLinearity
5         NewTree.LinearCoefficient = LinearCoefficient
6         NewTree.LinearValue = LinearValue
          
          Dim i As Long
7         For i = 1 To NumChildren
8             NewTree.SetChild i, Children(i)
9         Next i
          
10        Set Clone = NewTree

ExitFunction:
11        If RaiseError Then RethrowError
12        Exit Function

ErrorHandler:
13        If Not ReportError("ExpressionTree", "Clone") Then Resume
14        RaiseError = True
15        GoTo ExitFunction
End Function

' Sets the current ExpressionTree to be the same as another
Sub SetExpressionTree(NewTree As ExpressionTree)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         NodeText = NewTree.NodeText
4         NodeType = NewTree.NodeType
5         NodeLinearity = NewTree.NodeLinearity
6         LinearCoefficient = NewTree.LinearCoefficient
7         LinearValue = NewTree.LinearValue
          
8         ClearChildren
          
          Dim i As Long
9         For i = 1 To NewTree.GetNumChildren
10            SetChild i, NewTree.GetChild(i)
11        Next i

ExitSub:
12        If RaiseError Then RethrowError
13        Exit Sub

ErrorHandler:
14        If Not ReportError("ExpressionTree", "SetExpressionTree") Then Resume
15        RaiseError = True
16        GoTo ExitSub
End Sub

' Converts a linear tree into a Dictionary and Double
Public Sub ConvertLinearTreeToConstraint(constraint As Dictionary, constant As Double)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          ' Remove everything but "plus", variables, and constants
3         CollateLinearTerms
    
          ' Add linear terms into constraint/constant
4         AddLinearTreeToConstraint constraint, constant

ExitSub:
5         If RaiseError Then RethrowError
6         Exit Sub

ErrorHandler:
7         If Not ReportError("ExpressionTree", "ConvertLinearTreeToConstraint") Then Resume
8         RaiseError = True
9         GoTo ExitSub
End Sub

' Traverses a linear tree and replaces all operators other than "plus"
' We want a tree with constants, variables, and "plus" nodes only
Sub CollateLinearTerms()
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          ' First collate all children
          Dim i As Long
3         For i = 1 To NumChildren
4             Children(i).CollateLinearTerms
5         Next i

6         Select Case NodeLinearity
          ' If the node if constant, we want to evaluate it
          Case ExpressionTreeConstant
7             Select Case NodeType
              ' Constant expression on a number node - we read the number
              Case ExpressionTreeNumber
8                 LinearValue = Val(NodeText)
              
              ' Constant expression under operator - we can evaluate it and discard all children
9             Case ExpressionTreeOperator
10                EvaluateConstantOperator
11                LinearValue = Val(NodeText)
12            End Select
              
          ' A linear node - we need to update the coefficient on the linear term
13        Case ExpressionTreeLinear
14            Select Case NodeType
              ' Variable - Intitialise coefficient at 1 if hasn't been already
              Case ExpressionTreeVariable
15                If LinearCoefficient = 0 Then
16                    LinearCoefficient = 1
17                End If
              
              ' Linear operator node - we need to collate the constant and variable terms into a single variable node if possible
18            Case ExpressionTreeOperator
19                Select Case NodeText
                  
                  ' Do nothing on a plus node. This is just to escape the Else case
                  Case "plus"
                  
                  ' If we have a minus node, convert it to a plus and flip the sign on the second operator
20                Case "minus"
21                    Children(2).MultiplyByConstant -1
22                    NodeText = "plus"
                      
                  ' Negate the coefficient on the variable node and remove the "neg" node
23                Case "neg"
24                    Children(1).MultiplyByConstant -1
25                    SetTreeToChild 1
                          
                  ' There is a constant node and a linear node attached to this mult.
                  ' We multiply the constant through the linear node (which may be a "plus" tree)
26                Case "mult"
27                    If Children(1).NodeType = ExpressionTreeConstant Then
28                        Children(2).MultiplyByConstant Children(1).LinearValue
29                        SetTreeToChild 2
30                    Else ' Child 2 is the constant node
31                        Children(1).MultiplyByConstant Children(2).LinearValue
32                        SetTreeToChild 1
33                    End If
                  
                  ' First node is variable
34                Case "div"
35                    Children(1).MultiplyByConstant 1 / Children(2).LinearValue
36                    SetTreeToChild 1
                      
                  ' Second node must be 1 (as we know the expression is linear)
37                Case "pow"
38                    SetTreeToChild 1
          
39                Case Else
40                    RaiseGeneralError "Unknown linear operator " & NodeText & vbCrLf & "Please let us know about this so we can fix it."
41                End Select
42            End Select
43        End Select

ExitSub:
44        If RaiseError Then RethrowError
45        Exit Sub

ErrorHandler:
46        If Not ReportError("ExpressionTree", "CollateLinearTerms") Then Resume
47        RaiseError = True
48        GoTo ExitSub
End Sub

' Multiplies all nodes in the linear tree by a constant
Public Sub MultiplyByConstant(constant As Double)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         Select Case NodeType
          Case ExpressionTreeConstant
4             LinearValue = LinearValue * constant
5             NodeText = StrExNoPlus(LinearValue)
          
6         Case ExpressionTreeVariable
7             LinearCoefficient = LinearCoefficient * constant
              
8         Case ExpressionTreeOperator ' plus
9             Children(1).MultiplyByConstant constant
10            Children(2).MultiplyByConstant constant
              
11        End Select

ExitSub:
12        If RaiseError Then RethrowError
13        Exit Sub

ErrorHandler:
14        If Not ReportError("ExpressionTree", "MultiplyByConstant") Then Resume
15        RaiseError = True
16        GoTo ExitSub
End Sub

' Convert a collated linear tree into a Dictionary and Double constant
Sub AddLinearTreeToConstraint(constraint As Dictionary, constant As Double)
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          ' By this point, only "plus", variables and constants remain in the tree.
          ' We traverse the tree and add each variable we find to the constraint,
          ' and each constant value to the constant total
3         Select Case NodeType
          ' Add variable coefficient to constraint
          Case ExpressionTreeVariable
              Dim VarIndex As Long
4             VarIndex = VariableIndex.Item(NodeText)
5             constraint.Item(VarIndex) = constraint.Item(VarIndex) + LinearCoefficient
          
          ' Add constant
6         Case ExpressionTreeNumber
7             constant = constant + LinearValue
          
          ' Recurse on all "plus" nodes
8         Case ExpressionTreeOperator
9             If NodeText = "plus" Then
                  Dim i As Long
10                For i = 1 To NumChildren
11                    Children(i).AddLinearTreeToConstraint constraint, constant
12                Next i
13            Else
14                RaiseGeneralError "Operator that isn't plus at the last stage of the linear tree pruning: " & NodeText & vbCrLf & "Please let us know about this so we can fix it."
15            End If
16        End Select

ExitSub:
17        If RaiseError Then RethrowError
18        Exit Sub

ErrorHandler:
19        If Not ReportError("ExpressionTree", "AddLinearTreeToConstraint") Then Resume
20        RaiseError = True
21        GoTo ExitSub
End Sub

