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
    Set rngAddress = Nothing
    strAddress = ""
    
    lngDepth = 0
    boolCanEval = False
    boolIsConstant = False
    intRefsAdjCell = AdjCellUnknown
    boolEvaledWithCalculate = False
    
    Set Children = New Collection
    Set Parents = New Collection
End Sub


Public Function GetSheet() As String
    If rngAddress Is Nothing Then
        GetSheet = ""
    Else
        GetSheet = rngAddress.Parent.Name
    End If
End Function


Public Function IsDependentOn(rngAdjDepedents As Range) As Boolean
    IsDependentOn = True
    ' Check this just in case. Shouldn't ever happen, because
    ' there should be sheet check before this is done, and for
    ' that to happen it must NOT be Nothing.
    If rngAddress Is Nothing Then Exit Function
    'If Intersect(rngAddress, rngAdjDepedents) Is Nothing Then
    IsDependentOn = TestIntersect(rngAddress, rngAdjDepedents)
End Function

