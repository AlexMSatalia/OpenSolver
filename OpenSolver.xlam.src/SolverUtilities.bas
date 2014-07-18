Attribute VB_Name = "SolverUtilities"
' Functions that relate to multiple solvers, or delegate to internal solver functions.
Option Explicit
Public Const LPFileName As String = "model.lp"
Public Const XMLFileName As String = "model.ampl"
Public Const PuLPFileName As String = "opensolver.py"

Function SolverAvailable(Solver As String, Optional SolverPath As String, Optional errorString As String) As Boolean
' Delegated function returns True if solver is available and sets SolverPath to location of solver
    
    'All Neos solver always available
    If RunsOnNeos(Solver) Then
        SolverAvailable = True
        SolverPath = ""
        Exit Function
    End If

    Select Case Solver
    Case "CBC"
        SolverAvailable = SolverAvailable_CBC(SolverPath, errorString)
    Case "Gurobi"
        SolverAvailable = SolverAvailable_Gurobi(SolverPath, errorString)
    Case "NOMAD"
        SolverAvailable = SolverAvailable_NOMAD(errorString)
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
    Case "NeosCBC"
        SolverType = SolverType_NeosCBC
    Case "NOMAD"
        SolverType = SolverType_NOMAD
    Case "NeosBon"
        SolverType = SolverType_NeosBon
    Case "NeosCou"
        SolverType = SolverType_NeosCou
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

Function ReadModel(Solver As String, SolutionFilePathName As String, errorString As String) As Boolean
    Select Case Solver
    Case "CBC"
        ReadModel = ReadModel_CBC(SolutionFilePathName, errorString)
    Case "Gurobi"
        ReadModel = ReadModel_Gurobi(SolutionFilePathName, errorString)
    Case Else
        ReadModel = False
        errorString = "The solver " & Solver & " has not yet been incorporated fully into OpenSolver."
    End Select
End Function

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
    ModelFilePath = ModelFile(Solver)
    ' If model file is empty, then don't return anything
    If ModelFilePath <> "" Then
        ModelFilePath = GetTempFilePath(ModelFilePath)
    End If
End Function

Function RunsOnNeos(Solver As String) As Boolean
    If Solver Like "Neos*" Then
        RunsOnNeos = True
    Else
        RunsOnNeos = False
    End If
End Function

Sub GetNeosValues(Solver As String, Category As String, SolverType As String)
    Select Case Solver
    Case "NeosCBC"
        Category = "milp"
        SolverType = "cbc"
    Case "NeosBon"
        Category = "minco"
        SolverType = "Bonmin"
    Case "NeosCou"
        Category = "minco"
        SolverType = "Couenne"
    End Select
End Sub

Function UsesTokeniser(Solver As String) As Boolean
    Select Case Solver
    Case "PuLP", "NeosBon", "NeosCou"
        UsesTokeniser = True
    Case Else
        UsesTokeniser = False
    End Select
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
    Case "NOMAD"
        Set SolverNOMAD.OpenSolver_NOMAD = OpenSolver
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

Function SolverTitle(Solver As String) As String
    Select Case Solver
    Case "CBC"
        SolverTitle = SolverTitle_CBC
    Case "Gurobi"
        SolverTitle = SolverTitle_Gurobi
    Case "NOMAD"
        SolverTitle = SolverTitle_NOMAD
    Case "NeosCBC"
        SolverTitle = SolverTitle_NeosCBC
    Case "NeosCou"
        SolverTitle = SolverTitle_NeosCou
    Case "NeosBon"
        SolverTitle = SolverTitle_NeosBon
    End Select
End Function

Function ReverseSolverTitle(SolverTitle As String) As String
    Select Case SolverTitle
    Case SolverTitle_CBC
        ReverseSolverTitle = "CBC"
    Case SolverTitle_Gurobi
        ReverseSolverTitle = "Gurobi"
    Case SolverTitle_NOMAD
        ReverseSolverTitle = "NOMAD"
    Case SolverTitle_NeosCBC
        ReverseSolverTitle = "NeosCBC"
    Case SolverTitle_NeosCou
        ReverseSolverTitle = "NeosCou"
    Case SolverTitle_NeosBon
        ReverseSolverTitle = "NeosBon"
    End Select
End Function

Function SolverDesc(Solver As String) As String
    Select Case Solver
    Case "CBC"
        SolverDesc = SolverDesc_CBC
    Case "Gurobi"
        SolverDesc = SolverDesc_Gurobi
    Case "NOMAD"
        SolverDesc = SolverDesc_NOMAD
    Case "NeosCBC"
        SolverDesc = SolverDesc_NeosCBC
    Case "NeosCou"
        SolverDesc = SolverDesc_NeosCou
    Case "NeosBon"
        SolverDesc = SolverDesc_NeosBon
    End Select
End Function

Function SolverLink(Solver As String) As String
    Select Case Solver
    Case "CBC"
        SolverLink = SolverLink_CBC
    Case "Gurobi"
        SolverLink = SolverLink_Gurobi
    Case "NOMAD"
        SolverLink = SolverLink_NOMAD
    Case "NeosCBC"
        SolverLink = SolverLink_NeosCBC
    Case "NeosCou"
        SolverLink = SolverLink_NeosCou
    Case "NeosBon"
        SolverLink = SolverLink_NeosBon
    End Select
End Function

Function SolverHasSensitivityAnalysis(Solver As String) As Boolean
    ' Non-linear solvers don't have sensitivity analysis
    If SolverType(Solver) = OpenSolver_SolverType.NonLinear Then
        SolverHasSensitivityAnalysis = False
    End If
    
    Select Case Solver
    Case "CBC", "Gurobi"
        SolverHasSensitivityAnalysis = True
    Case Else
        SolverHasSensitivityAnalysis = False
    End Select
End Function
