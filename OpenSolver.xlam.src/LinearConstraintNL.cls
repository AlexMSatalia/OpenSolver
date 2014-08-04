VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinearConstraintNL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim VariablesPresent() As Boolean
Dim Coefficients() As Double
Dim NumVariables As Integer

Public Property Let Count(numVars As Integer)
    NumVariables = numVars
    ReDim VariablesPresent(NumVariables)
    ReDim Coefficients(NumVariables)
End Property

Public Property Get Count() As Integer
    Count = NumVariables
End Property

Public Property Let VariablePresent(index As Integer, Present As Boolean)
    VariablesPresent(index) = Present
End Property

Public Property Get VariablePresent(index As Integer) As Boolean
    VariablePresent = VariablesPresent(index)
End Property

Public Property Let Coefficient(index As Integer, value As Double)
    Coefficients(index) = value
End Property

Public Property Get Coefficient(index As Integer) As Double
    Coefficient = Coefficients(index)
End Property

Public Property Get NumPresent() As Integer
    NumPresent = 0
    
    Dim i As Integer
    For i = 1 To NumVariables
        If VariablesPresent(i) Then
            NumPresent = NumPresent + 1
        End If
    Next i
End Property

Public Function Display() As String
    Dim i As Integer
    For i = 1 To NumVariables
        Display = Display & Coefficients(i) & " "
    Next i
End Function

Public Sub InvertCoefficients()
    Dim i As Integer
    For i = 1 To NumVariables
        Coefficients(i) = -Coefficients(i)
    Next i
End Sub
