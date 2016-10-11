VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Location
Public rngAddress As Range
Public strAddress As String
' Formula
Public strFormula As String
Public strFormulaParsed As String
' Initial value
Public initialValue As Double
' Properties
Public lngDepth As Long
Public boolIsRoot As Boolean
Public boolCanEval As Boolean
Public boolIsConstant As Boolean
Public Enum AdjCellStatus
    AdjCellIndependent
    AdjCellDependent
    AdjCellUnknown
End Enum
Public intRefsAdjCell As AdjCellStatus
Public boolEvaledWithCalculate As Boolean
Public boolCanBeRemoved As Boolean

' Graph structure
Public Children As Collection
Public Parents As Collection
 
Private Sub Class_Initialize()
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         Set rngAddress = Nothing
4         strAddress = ""
          
5         lngDepth = 0
6         boolCanEval = False
7         boolIsConstant = False
8         intRefsAdjCell = AdjCellUnknown
9         boolEvaledWithCalculate = False
          
10        Set Children = New Collection
11        Set Parents = New Collection

ExitSub:
12        If RaiseError Then RethrowError
13        Exit Sub

ErrorHandler:
14        If Not ReportError("CFormula", "Class_Initialize") Then Resume
15        RaiseError = True
16        GoTo ExitSub
End Sub


Public Function GetSheet() As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         If rngAddress Is Nothing Then
4             GetSheet = ""
5         Else
6             GetSheet = rngAddress.Parent.Name
7         End If

ExitFunction:
8         If RaiseError Then RethrowError
9         Exit Function

ErrorHandler:
10        If Not ReportError("CFormula", "GetSheet") Then Resume
11        RaiseError = True
12        GoTo ExitFunction
End Function

Public Function IsDependentOn(rngAdjDepedents As Range) As Boolean
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

3         IsDependentOn = True
          ' Check this just in case. Shouldn't ever happen, because
          ' there should be sheet check before this is done, and for
          ' that to happen it must NOT be Nothing.
4         If rngAddress Is Nothing Then GoTo ExitFunction
          'If Intersect(rngAddress, rngAdjDepedents) Is Nothing Then
5         IsDependentOn = TestIntersect(rngAddress, rngAdjDepedents)

ExitFunction:
6         If RaiseError Then RethrowError
7         Exit Function

ErrorHandler:
8         If Not ReportError("CFormula", "IsDependentOn") Then Resume
9         RaiseError = True
10        GoTo ExitFunction
End Function

