VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExpressionTreeStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' Stack for ExpressionTree objects
'==============================================================================
Option Explicit

Private colTreeStack As Collection

Private Sub Class_Initialize()
1         Set colTreeStack = New Collection
End Sub

Public Sub Push(Tree As ExpressionTree)
1         colTreeStack.Add Tree
End Sub

Public Function Pop() As ExpressionTree
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim obj As ExpressionTree
3         Set obj = colTreeStack(colTreeStack.Count)
4         colTreeStack.Remove (colTreeStack.Count)
5         Set Pop = obj

ExitFunction:
6         If RaiseError Then RethrowError
7         Exit Function

ErrorHandler:
8         If Not ReportError("ExpressionTreeStack", "Pop") Then Resume
9         RaiseError = True
10        GoTo ExitFunction
End Function

Public Function Peek() As ExpressionTree
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim obj As ExpressionTree
3         Set obj = colTreeStack(colTreeStack.Count)
4         Set Peek = obj

ExitFunction:
5         If RaiseError Then RethrowError
6         Exit Function

ErrorHandler:
7         If Not ReportError("ExpressionTreeStack", "Peek") Then Resume
8         RaiseError = True
9         GoTo ExitFunction
End Function

Property Get Count() As Long
1         Count = colTreeStack.Count
End Property

Public Function Display() As String
          Dim i As Long, output As String
1         For i = 1 To colTreeStack.Count
2             output = output & colTreeStack(i).Display & ", "
3         Next i
4         Display = output
End Function


