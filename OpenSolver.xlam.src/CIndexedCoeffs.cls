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
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If NumEntries = 0 Then
4             ReDim Indices(10) As Long
5             ReDim Values(10) As Double
6         ElseIf NumEntries = UBound(Indices) Then
              ' Array needs to grow
7             ReDim Preserve Indices(UBound(Indices) + Int(UBound(Indices) * 0.5) + 1) As Long
8             ReDim Preserve Values(UBound(Indices)) As Double
9         End If
10        NumEntries = NumEntries + 1
11        Indices(NumEntries) = Index
12        Values(NumEntries) = value

ExitSub:
13        If RaiseError Then RethrowError
14        Exit Sub

ErrorHandler:
15        If Not ReportError("CIndexedCoeffs", "Add") Then Resume
16        RaiseError = True
17        GoTo ExitSub
End Sub

Public Function Index(i As Long) As Long
1         Index = Indices(i)
End Function

Public Function Coefficient(i As Long) As Double
1         Coefficient = Values(i)
End Function

Public Function Count() As Long
1         Count = NumEntries
End Function

Public Function Evaluate(DecisionVariableValues() As Double) As Double
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          ' Evaluate the equation's value for the given variable values
          Dim i As Long, v As Double
3         For i = 1 To NumEntries
4             v = v + Coefficient(i) * DecisionVariableValues(Index(i))
5         Next i
6         Evaluate = v

ExitFunction:
7         If RaiseError Then RethrowError
8         Exit Function

ErrorHandler:
9         If Not ReportError("CIndexedCoeffs", "Evaluate") Then Resume
10        RaiseError = True
11        GoTo ExitFunction
End Function

Public Function Evaluate_RecordPrecision(DecisionVariableValues() As Double, vMax As Double) As Double
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          ' Evaluate the equation's value for the given variable values
          Dim i As Long, v As Double, value As Double, coeff As Double
3         vMax = 0
4         For i = 1 To NumEntries
5             coeff = Coefficient(i)
6             value = DecisionVariableValues(Index(i))
7             v = v + coeff * value
8             If Abs(v) > vMax Then vMax = Abs(v)
9             If Abs(coeff) > vMax Then vMax = Abs(coeff)
10            If Abs(value) > vMax Then vMax = Abs(value)
11        Next i
12        Evaluate_RecordPrecision = v

ExitFunction:
13        If RaiseError Then RethrowError
14        Exit Function

ErrorHandler:
15        If Not ReportError("CIndexedCoeffs", "Evaluate_RecordPrecision") Then Resume
16        RaiseError = True
17        GoTo ExitFunction
End Function

Public Function Clone() As CIndexedCoeffs
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          ' Return a copy of this class
          Dim v As CIndexedCoeffs, i As Long
3         Set v = New CIndexedCoeffs
          'ReDim v.Indices(count)
          'ReDim v.Values(count)
4         For i = 1 To Count
5             v.Add Index(i), Coefficient(i)
6         Next i
7         Set Clone = v

ExitFunction:
8         If RaiseError Then RethrowError
9         Exit Function

ErrorHandler:
10        If Not ReportError("CIndexedCoeffs", "Clone") Then Resume
11        RaiseError = True
12        GoTo ExitFunction
End Function

Public Function AsVector(Length As Long) As Double()
          Dim coeffs() As Double
1         ReDim coeffs(1 To Length) As Double
          
          ' Zero it out to be sure
          Dim i As Long
2         For i = 1 To Length
3             coeffs(i) = 0
4         Next i
          
          ' Add in the results of the sparse vector
5         For i = 1 To Count
6             coeffs(Index(i)) = Coefficient(i)
7         Next i
          
8         AsVector = coeffs()
End Function
