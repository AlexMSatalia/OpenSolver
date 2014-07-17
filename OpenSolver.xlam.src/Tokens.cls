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
    Set colTokens = New Collection
End Sub

Public Property Get NewEnum() As IUnknown
    Set NewEnum = colTokens.[_NewEnum]
End Property

Public Sub Add(Token As Token, Optional key As Variant)
    colTokens.Add Token, key
    Set Token.Parent = Me
End Sub

Public Sub Remove(index As Variant)
    colTokens.Remove index
End Sub

Function Item(index As Variant) As Token
    Set Item = colTokens.Item(index)
End Function

Property Get Count() As Long
    Count = colTokens.Count
End Property

Public Sub Clear()
    Set colTokens = New Collection
End Sub

Public Function NewToken(Text As String, TokenType As Long, FormulaIndex As Long, FormulaLength As Long) As Token
    Dim obj As Token

    Set obj = New Token
    obj.Text = Text
    obj.TokenType = TokenType
    obj.FormulaIndex = FormulaIndex
    obj.FormulaLength = FormulaLength

    Set NewToken = obj
End Function

