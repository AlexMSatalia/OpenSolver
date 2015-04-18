VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Interface common to all solvers

Public Property Get Name() As String
    ' Human-readable name of the solver being used
    ' NEOS solvers should use the actual name of the solver
End Property

Public Property Get ShortName() As String
    ' Short code for the solver
    ' Should match the names in GetAvailableSolvers()
End Property

Public Property Get Title() As String
    ' Short title string for solver change form
    ' Solver name and type
End Property

Public Property Get Desc() As String
    ' Description for solver change form
End Property

Public Property Get Link() As String
    ' Hyperlink for solver change form
End Property

Public Property Get IterationLimitName() As String
    ' Name of parameter controlling iteration limit
    ' Leave empty to mark unavailable
End Property

Public Property Get PrecisionName() As String
    ' Name of parameter controlling precision
    ' Leave empty to mark unavailable
End Property

Public Property Get TimeLimitName() As String
    ' Name of parameter controlling time limit
    ' Leave empty to mark unavailable
End Property

Public Property Get ToleranceName() As String
    ' Name of parameter controlling tolerance
    ' Leave empty to mark unavailable
End Property

Public Property Get ModelType() As OpenSolver_ModelType
End Property

Public Sub CleanFiles()
End Sub

Public Sub CheckLog(s As COpenSolver)
End Sub
