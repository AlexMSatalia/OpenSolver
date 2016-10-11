VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TokenStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' OpenSolver
' Formula tokenizer functionality is from http://www.dailydoseofexcel.com
' Code is written by Rob van Gelder
' http://www.dailydoseofexcel.com/archives/2009/12/05/formula-tokenizer/
' GUI-related functionality stripped.
'==============================================================================
Option Explicit

Private colTokens As Collection

Private Sub Class_Initialize()
1         Set colTokens = New Collection
End Sub

Public Sub Push(Token As Token)
1         colTokens.Add Token
End Sub

Public Function Pop() As Token
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim obj As Token
3         Set obj = colTokens(colTokens.Count)
4         colTokens.Remove (colTokens.Count)
5         Set Pop = obj

ExitFunction:
6         If RaiseError Then RethrowError
7         Exit Function

ErrorHandler:
8         If Not ReportError("TokenStack", "Pop") Then Resume
9         RaiseError = True
10        GoTo ExitFunction
End Function

Public Function Peek() As Token
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim obj As Token
3         Set obj = colTokens(colTokens.Count)
4         Set Peek = obj

ExitFunction:
5         If RaiseError Then RethrowError
6         Exit Function

ErrorHandler:
7         If Not ReportError("TokenStack", "Peek") Then Resume
8         RaiseError = True
9         GoTo ExitFunction
End Function

Property Get Count() As Long
1         Count = colTokens.Count
End Property

