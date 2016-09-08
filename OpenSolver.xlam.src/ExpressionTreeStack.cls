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
8498      Set colTreeStack = New Collection
End Sub

Public Sub Push(Tree As ExpressionTree)
8499      colTreeStack.Add Tree
End Sub

Public Function Pop() As ExpressionTree
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim obj As ExpressionTree
8500      Set obj = colTreeStack(colTreeStack.Count)
8501      colTreeStack.Remove (colTreeStack.Count)
8502      Set Pop = obj

ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ExpressionTreeStack", "Pop") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Public Function Peek() As ExpressionTree
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim obj As ExpressionTree
8503      Set obj = colTreeStack(colTreeStack.Count)
8504      Set Peek = obj

ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("ExpressionTreeStack", "Peek") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Property Get Count() As Long
8505      Count = colTreeStack.Count
End Property

Public Function Display() As String
          Dim i As Long, output As String
8506      For i = 1 To colTreeStack.Count
8507          output = output & colTreeStack(i).Display & ", "
8508      Next i
8509      Display = output
End Function


