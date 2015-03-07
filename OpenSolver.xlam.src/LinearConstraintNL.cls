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
8992      NumVariables = numVars
8993      ReDim VariablesPresent(NumVariables)
8994      ReDim Coefficients(NumVariables)
End Property

Public Property Get Count() As Long
8995      Count = NumVariables
End Property

Public Property Let VariablePresent(Index As Long, Present As Boolean)
8996      VariablesPresent(Index) = Present
End Property

Public Property Get VariablePresent(Index As Long) As Boolean
8997      VariablePresent = VariablesPresent(Index)
End Property

Public Property Let Coefficient(Index As Long, value As Double)
8998      Coefficients(Index) = value
End Property

Public Property Get Coefficient(Index As Long) As Double
8999      Coefficient = Coefficients(Index)
End Property

' The number of variables present in the constraint
Public Property Get NumPresent() As Long
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

9000      NumPresent = 0
          
          Dim i As Long
9001      For i = 1 To NumVariables
9002          If VariablesPresent(i) Then
9003              NumPresent = NumPresent + 1
9004          End If
9005      Next i

ExitProperty:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Property

ErrorHandler:
          If Not ReportError("LinearConstraintNL", "NumPresent") Then Resume
          RaiseError = True
          GoTo ExitProperty
End Property

' Outputs all variable coefficients to string for debugging
Public Function Display() As String
          Dim i As Long
9006      For i = 1 To NumVariables
9007          Display = Display & Coefficients(i) & " "
9008      Next i
End Function

' Inverts the sign on all coefficients. Used for negating the expression
Public Sub InvertCoefficients()
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim i As Long
9009      For i = 1 To NumVariables
9010          If Coefficients(i) <> 0 Then
9011              Coefficients(i) = -Coefficients(i)
9012          End If
9013      Next i

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("LinearConstraintNL", "InvertCoefficients") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Public Function IsEmpty() As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

9014      IsEmpty = True
          Dim i As Long
9015      For i = 1 To NumVariables
9016          If Coefficients(i) <> 0 Then
9017              IsEmpty = False
9018              GoTo ExitFunction
9019          End If
9020      Next i

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("LinearConstraintNL", "IsEmpty") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function
