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

' Class for storing information about a linear expression (constraint/objective)

Dim VariablesPresent() As Boolean   ' Whether each variable is present in the expression
Dim Coefficients() As Double        ' The coefficient of each variable if present
Dim NumVariables As Long         ' The number of variables present

Public Property Let Count(numVars As Long)
    NumVariables = numVars
    ReDim VariablesPresent(NumVariables)
    ReDim Coefficients(NumVariables)
End Property

Public Property Get Count() As Long
    Count = NumVariables
End Property

Public Property Let VariablePresent(index As Long, Present As Boolean)
    VariablesPresent(index) = Present
End Property

Public Property Get VariablePresent(index As Long) As Boolean
    VariablePresent = VariablesPresent(index)
End Property

Public Property Let Coefficient(index As Long, value As Double)
    Coefficients(index) = value
End Property

Public Property Get Coefficient(index As Long) As Double
    Coefficient = Coefficients(index)
End Property

' The number of variables present in the constraint
Public Property Get NumPresent() As Long
    NumPresent = 0
    
    Dim i As Long
    For i = 1 To NumVariables
        If VariablesPresent(i) Then
            NumPresent = NumPresent + 1
        End If
    Next i
End Property

' Outputs all variable coefficients to string for debugging
Public Function Display() As String
    Dim i As Long
    For i = 1 To NumVariables
        Display = Display & Coefficients(i) & " "
    Next i
End Function

' Inverts the sign on all coefficients. Used for negating the expression
Public Sub InvertCoefficients()
    Dim i As Long
    For i = 1 To NumVariables
        If Coefficients(i) <> 0 Then
            Coefficients(i) = -Coefficients(i)
        End If
    Next i
End Sub

Public Function IsEmpty() As Boolean
    IsEmpty = True
    Dim i As Long
    For i = 1 To NumVariables
        If Coefficients(i) <> 0 Then
            IsEmpty = False
            Exit Function
        End If
    Next i
End Function
