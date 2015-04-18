VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIndexedCoeffs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' This implements a sparse row representation used in our sparse arrays

Private Indices() As Long
Private Values() As Double

Private NumEntries As Long

Public Sub Add(Index As Long, value As Double)
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

3550      If NumEntries = 0 Then
3551          ReDim Indices(10) As Long
3552          ReDim Values(10) As Double
3553      ElseIf NumEntries = UBound(Indices) Then
              ' Array needs to grow
3554          ReDim Preserve Indices(UBound(Indices) + Int(UBound(Indices) * 0.5) + 1) As Long
3555          ReDim Preserve Values(UBound(Indices)) As Double
3556      End If
3557      NumEntries = NumEntries + 1
3558      Indices(NumEntries) = Index
3559      Values(NumEntries) = value

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("CIndexedCoeffs", "Add") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Public Function Index(i As Long) As Long
3560      Index = Indices(i)
End Function

Public Function Coefficient(i As Long) As Double
3561      Coefficient = Values(i)
End Function

Public Function Count() As Long
3562      Count = NumEntries
End Function

Public Function Evaluate(DecisionVariableValues() As Double) As Double
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          ' Evaluate the equation's value for the given variable values
          Dim i As Long, v As Double
3563      For i = 1 To NumEntries
3564          v = v + Coefficient(i) * DecisionVariableValues(Index(i))
3566      Next i
3567      Evaluate = v

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("CIndexedCoeffs", "Evaluate") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Public Function Evaluate_RecordPrecision(DecisionVariableValues() As Double, vMax As Double) As Double
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          ' Evaluate the equation's value for the given variable values
          Dim i As Long, v As Double, value As Double, coeff As Double
3568      vMax = 0
3569      For i = 1 To NumEntries
3570          coeff = Coefficient(i)
3571          value = DecisionVariableValues(Index(i))
3572          v = v + coeff * value
3573          If Abs(v) > vMax Then vMax = Abs(v)
3574          If Abs(coeff) > vMax Then vMax = Abs(coeff)
3575          If Abs(value) > vMax Then vMax = Abs(value)
3576      Next i
3577      Evaluate_RecordPrecision = v

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("CIndexedCoeffs", "Evaluate_RecordPrecision") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Public Function Clone() As CIndexedCoeffs
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          ' Return a copy of this class
          Dim v As CIndexedCoeffs, i As Long
3578      Set v = New CIndexedCoeffs
          'ReDim v.Indices(count)
          'ReDim v.Values(count)
3579      For i = 1 To Count
3580          v.Add Index(i), Coefficient(i)
3581      Next i
3582      Set Clone = v

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("CIndexedCoeffs", "Clone") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function
