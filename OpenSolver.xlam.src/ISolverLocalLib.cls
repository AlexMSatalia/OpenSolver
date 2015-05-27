VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISolverLocalLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetLibPath(Optional errorString As String, Optional Bitness As String) As String
End Function

Public Property Get LibName() As String
End Property

Public Property Get LibBinary() As String
End Property

Public Property Get LibVersion() As String
End Property

Public Sub Solve(s As COpenSolver)
End Sub
