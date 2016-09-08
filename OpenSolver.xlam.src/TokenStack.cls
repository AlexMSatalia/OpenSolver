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
5482      Set colTokens = New Collection
End Sub

Public Sub Push(Token As Token)
5483      colTokens.Add Token
End Sub

Public Function Pop() As Token
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim obj As Token
5484      Set obj = colTokens(colTokens.Count)
5485      colTokens.Remove (colTokens.Count)
5486      Set Pop = obj

ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("TokenStack", "Pop") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Public Function Peek() As Token
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim obj As Token
5487      Set obj = colTokens(colTokens.Count)
5488      Set Peek = obj

ExitFunction:
          If RaiseError Then RethrowError
          Exit Function

ErrorHandler:
          If Not ReportError("TokenStack", "Peek") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Property Get Count() As Long
5489      Count = colTokens.Count
End Property

