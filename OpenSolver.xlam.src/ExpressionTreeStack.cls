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
    Set colTreeStack = New Collection
End Sub

Public Sub Push(Tree As ExpressionTree)
    colTreeStack.Add Tree
End Sub

Public Function Pop() As ExpressionTree
    Dim obj As ExpressionTree
    Set obj = colTreeStack(colTreeStack.Count)
    colTreeStack.Remove (colTreeStack.Count)
    Set Pop = obj
End Function

Public Function Peek() As ExpressionTree
    Dim obj As ExpressionTree
    Set obj = colTreeStack(colTreeStack.Count)
    Set Peek = obj
End Function

Property Get Count() As Long
    Count = colTreeStack.Count
End Property

Public Function Display() As String
    Dim i As Long, output As String
    For i = 1 To colTreeStack.Count
        output = output & colTreeStack(i).Display & ", "
    Next i
    Display = output
End Function


