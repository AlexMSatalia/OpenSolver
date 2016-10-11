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
1         Set colTokens = New Collection
End Sub

Public Property Get NewEnum() As IUnknown
1         Set NewEnum = colTokens.[_NewEnum]
End Property

Public Sub Add(Token As Token, Optional Key As Variant)
1         colTokens.Add Token, Key
2         Set Token.Parent = Me
End Sub

Public Sub Remove(Index As Variant)
1         colTokens.Remove Index
End Sub

Function Item(Index As Variant) As Token
1         Set Item = colTokens.Item(Index)
End Function

Property Get Count() As Long
1         Count = colTokens.Count
End Property

Public Sub Clear()
1         Set colTokens = New Collection
End Sub

Public Function NewToken(Text As String, TokenType As Long, FormulaIndex As Long, FormulaLength As Long) As Token
          Dim obj As Token

1         Set obj = New Token
2         obj.Text = Text
3         obj.TokenType = TokenType
4         obj.FormulaIndex = FormulaIndex
5         obj.FormulaLength = FormulaLength

6         Set NewToken = obj
End Function

