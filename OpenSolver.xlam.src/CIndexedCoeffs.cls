VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIndexedCoeffs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' This implements a sparse row representation used in our sparse arrays

Private Indices() As Long
Private Values() As Double

Private NumEntries As Long

Public Sub Add(index As Long, Value As Double)
36360     If NumEntries = 0 Then
36370         ReDim Indices(10) As Long
36380         ReDim Values(10) As Double
36390     ElseIf NumEntries = UBound(Indices) Then
              ' Array needs to grow
36400         ReDim Preserve Indices(UBound(Indices) + Int(UBound(Indices) * 0.5) + 1) As Long
36410         ReDim Preserve Values(UBound(Indices)) As Double
36420     End If
36430     NumEntries = NumEntries + 1
36440     Indices(NumEntries) = index
36450     Values(NumEntries) = Value
End Sub

Public Function index(i As Long) As Long
36460     index = Indices(i)
End Function

Public Function Coefficient(i As Long) As Double
36470     Coefficient = Values(i)
End Function

Public Function count() As Long
36480     count = NumEntries
End Function

Public Function Evaluate(DecisionVariableValues() As Double) As Double
          ' Evaluate the equation's value for the given variable values
          Dim i As Long, v As Double, vMax As Double
36490     For i = 1 To NumEntries
36500         v = v + Coefficient(i) * DecisionVariableValues(index(i))
36510         If Abs(v) > vMax Then vMax = Abs(v)
36520     Next i
36530     Evaluate = v
End Function

Public Function Evaluate_RecordPrecision(DecisionVariableValues() As Double, vMax As Double) As Double
          ' Evaluate the equation's value for the given variable values
          Dim i As Long, v As Double, Value As Double, coeff As Double
36540     vMax = 0
36550     For i = 1 To NumEntries
36560         coeff = Coefficient(i)
36570         Value = DecisionVariableValues(index(i))
36580         v = v + coeff * Value
36590         If Abs(v) > vMax Then vMax = Abs(v)
36600         If Abs(coeff) > vMax Then vMax = Abs(coeff)
36610         If Abs(Value) > vMax Then vMax = Abs(Value)
36620     Next i
36630     Evaluate_RecordPrecision = v
End Function

Public Function Clone() As CIndexedCoeffs
          ' Return a copy of this class
          Dim v As CIndexedCoeffs, i As Long
36640     Set v = New CIndexedCoeffs
          'ReDim v.Indices(count)
          'ReDim v.Values(count)
36650     For i = 1 To count
36660         v.Add index(i), Coefficient(i)
36670     Next i
36680     Set Clone = v
End Function
