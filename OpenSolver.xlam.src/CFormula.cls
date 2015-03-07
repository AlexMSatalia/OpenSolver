VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
          RaiseError = False
          On Error GoTo ErrorHandler

4765      Set rngAddress = Nothing
4766      strAddress = ""
          
4767      lngDepth = 0
4768      boolCanEval = False
4769      boolIsConstant = False
4770      intRefsAdjCell = AdjCellUnknown
4771      boolEvaledWithCalculate = False
          
4772      Set Children = New Collection
4773      Set Parents = New Collection

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("CFormula", "Class_Initialize") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub


Public Function GetSheet() As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

4774      If rngAddress Is Nothing Then
4775          GetSheet = ""
4776      Else
4777          GetSheet = rngAddress.Parent.Name
4778      End If

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("CFormula", "GetSheet") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function


Public Function IsDependentOn(rngAdjDepedents As Range) As Boolean
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

4779      IsDependentOn = True
          ' Check this just in case. Shouldn't ever happen, because
          ' there should be sheet check before this is done, and for
          ' that to happen it must NOT be Nothing.
4780      If rngAddress Is Nothing Then GoTo ExitFunction
          'If Intersect(rngAddress, rngAdjDepedents) Is Nothing Then
4781      IsDependentOn = TestIntersect(rngAddress, rngAdjDepedents)

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("CFormula", "IsDependentOn") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

