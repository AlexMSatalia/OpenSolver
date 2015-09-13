VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISolverLocalExec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetExecPath(Optional errorString As String, Optional Bitness As String) As String
End Function

Public Property Get ExecName() As String
End Property

Public Function CreateSolveScript(ModelFilePathName As String, c As COpenSolver) As String
End Function

