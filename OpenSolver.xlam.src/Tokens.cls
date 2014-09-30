VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tokens"
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
5468      Set colTokens = New Collection
End Sub

Public Property Get NewEnum() As IUnknown
5469      Set NewEnum = colTokens.[_NewEnum]
End Property

Public Sub Add(Token As Token, Optional key As Variant)
5470      colTokens.Add Token, key
5471      Set Token.Parent = Me
End Sub

Public Sub Remove(index As Variant)
5472      colTokens.Remove index
End Sub

Function Item(index As Variant) As Token
5473      Set Item = colTokens.Item(index)
End Function

Property Get Count() As Long
5474      Count = colTokens.Count
End Property

Public Sub Clear()
5475      Set colTokens = New Collection
End Sub

Public Function NewToken(Text As String, TokenType As Long, FormulaIndex As Long, FormulaLength As Long) As Token
          Dim obj As Token

5476      Set obj = New Token
5477      obj.Text = Text
5478      obj.TokenType = TokenType
5479      obj.FormulaIndex = FormulaIndex
5480      obj.FormulaLength = FormulaLength

5481      Set NewToken = obj
End Function

