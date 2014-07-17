Attribute VB_Name = "SolverUtilities"
' Functions that relate to multiple solvers, or delegate to internal solver functions.
Option Explicit
Public Const LPFileName As String = "model.lp"
Public Const XMLFileName As String = "job.xml"
Public Const PuLPFileName As String = "opensolver.py"

Function SolverAvailable(Solver As String, ByRef SolverPath As String) As Boolean
' Delegated function returns True if solver is available and sets SolverPath to location of solver
    
    'All Neos solver always available
    If RunsOnNeos(Solver) Then
        SolverAvailable = True
        SolverPath = ""
        Exit Function
    End If

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

Function SolutionFilePath(Solver As String) As String
    Select Case Solver
    Case "CBC"
        SolutionFilePath = SolutionFilePath_CBC
    Case "Gurobi"
        SolutionFilePath = SolutionFilePath_Gurobi
    Case Else
        SolutionFilePath = ""
    End Select
End Function

Sub CleanFiles(errorPrefix)
    CleanFiles_CBC (errorPrefix)
    CleanFiles_Gurobi (errorPrefix)
End Sub

Function ModelFile(Solver As String) As String
    Select Case Solver
    Case "CBC", "Gurobi"
        ' output the model to an LP format text file
        ' See http://lpsolve.sourceforge.net/5.5/CPLEX-format.htm
        ModelFile = LPFileName
    Case "NeosCBC", "NeosCou", "NeosBon"
        ModelFile = XMLFileName
    Case "PuLP"
        ModelFile = PuLPFileName
    Case Else
        ModelFile = ""
    End Select
End Function

Function ModelFilePath(Solver As String) As String
    ModelFilePath = GetTempFilePath(ModelFile(Solver))
End Function

Function RunsOnNeos(Solver As String) As Boolean
    If Solver Like "Neos*" Then
        RunsOnNeos = True
    Else
        RunsOnNeos = False
    End If
End Function

Function GetExtraParameters(Solver As String, sheet As Worksheet, errorString As String) As String
    Select Case Solver
    Case "CBC"
        GetExtraParameters = GetExtraParameters_CBC(sheet, errorString)
    Case Else
        GetExtraParameters = ""
    End Select
End Function

Sub SetOpenSolver(Solver As String, OpenSolver As COpenSolver, SparseA() As CIndexedCoeffs)
    Select Case Solver
    Case "CBC"
        Set SolverCBC.OpenSolver_CBC = OpenSolver
        SolverCBC.SparseA_CBC = SparseA
    Case "Gurobi"
        Set SolverGurobi.OpenSolver_Gurobi = OpenSolver
        SolverGurobi.SparseA_Gurobi = SparseA
    End Select
End Sub

Function CreateSolveScript(Solver As String, SolutionFilePathName As String, ExtraParametersString As String, SolveOptions As SolveOptionsType) As String
    Select Case Solver
    Case "CBC"
        CreateSolveScript = CreateSolveScript_CBC(SolutionFilePathName, ExtraParametersString, SolveOptions)
    Case "Gurobi"
        CreateSolveScript = CreateSolveScript_Gurobi(SolutionFilePathName, ExtraParametersString, SolveOptions)
    End Select
End Function
