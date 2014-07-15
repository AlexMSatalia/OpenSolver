Attribute VB_Name = "SolverUtilities"
' Functions that relate to multiple solvers, or delegate to internal solver functions.
Option Explicit

Function SolverAvailable(Solver As String, ByRef SolverPath As String) As Boolean
' Delegated function returns True if solver is available and sets SolverPath to location of solver
    Select Case Solver
    Case "CBC"
        SolverAvailable = SolverAvailable_CBC(SolverPath)
    Case "Gurobi"
        SolverAvailable = SolverAvailable_Gurobi(SolverPath)
    Case Else
        SolverAvailable = False
        SolverPath = ""
    End Select
End Function

Function SolverType(Solver As String) As String
    Select Case Solver
    Case "CBC"
        SolverType = SolverType_CBC
    Case "Gurobi"
        SolverType = SolverType_Gurobi
    Case Else
        SolverType = OpenSolver_SolverType.Unknown
    End Select
End Function

Sub CleanFiles(errorPrefix)
    CleanFiles_CBC (errorPrefix)
    CleanFiles_Gurobi (errorPrefix)
End Sub
