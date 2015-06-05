VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSolverGurobi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISolver
Implements ISolverLocal
Implements ISolverLocalExec
Implements ISolverLinear
Implements ISolverFile

Private Const pTitle = "Gurobi (Linear solver)"
Private Const pDesc = "Gurobi is a solver for linear programming (LP), quadratic and quadratically constrained programming (QP and QCP), and mixed-integer programming (MILP, MIQP, and MIQCP). It requires the user to download and install a version of the Gurobi and to have GurobiOSRun.py in the OpenSolver directory."
Private Const pLink = "http://www.gurobi.com/resources/documentation"
Private Const pName = "Gurobi"
Private Const pShortName = "Gurobi"

Private Const pExecName = "gurobi_cl" & ExecExtension
Private Const pShellScript = "gurobi" & ScriptExtension
Private Const pScriptFile = "gurobi_tmp" & ScriptExtension
Private Const pPythonScript = "gurobiOSRun.py"

Private Const pSensitivityAnalysisAvailable = True

Private Const pSolutionFile = "modelsolution.sol"
Private Const pSensitivityFile = "sensitivityData.sol"
Private Const pFileType = OpenSolver_FileType.LP
Private Const pModelType = OpenSolver_ModelType.Diff

Private Const pIterationLimitName = "IterationLimit"
Private Const pTimeLimitName = "TimeLimit"
Private Const pToleranceName = "MIPGap"

Private Enum GurobiResult
    Optimal = 2
    Infeasible = 3
    InfOrUnbound = 4
    Unbounded = 5
    SolveStoppedIter = 7
    SolveStoppedTime = 9
    SolveStoppedUser = 11
    Unsolved = 12
    SubOptimal = 13
End Enum

' ISolver

Private Property Get ISolver_Name() As String
    ISolver_Name = pName
End Property

Private Property Get ISolver_ShortName() As String
    ISolver_ShortName = pShortName
End Property

Private Property Get ISolver_Title() As String
    ISolver_Title = pTitle
End Property

Private Property Get ISolver_Desc() As String
    ISolver_Desc = pDesc
End Property

Private Property Get ISolver_Link() As String
    ISolver_Link = pLink
End Property

Private Property Get ISolver_IterationLimitName() As String
    ISolver_IterationLimitName = pIterationLimitName
End Property

Private Property Get ISolver_PrecisionName() As String
End Property

Private Property Get ISolver_TimeLimitName() As String
    ISolver_TimeLimitName = pTimeLimitName
End Property

Private Property Get ISolver_ToleranceName() As String
    ISolver_ToleranceName = pToleranceName
End Property

Private Property Get ISolver_ModelType() As OpenSolver_ModelType
    ISolver_ModelType = pModelType
End Property

Private Sub ISolver_CleanFiles()
    CleanFiles
End Sub

Private Sub ISolver_CheckLog(s As COpenSolver)
    CheckLog s
End Sub

' ISolverLocal

Private Property Get ISolverLocal_Version() As String
    ISolverLocal_Version = Version()
End Property

Private Property Get ISolverLocal_Bitness() As String
    ISolverLocal_Bitness = Bitness()
End Property

' ISolverLocalExec

Private Property Get ISolverLocalExec_ExecName() As String
    ISolverLocalExec_ExecName = pExecName
End Property

Private Function ISolverLocalExec_GetExecPath(Optional errorString As String, Optional Bitness As String) As String
    IsPresent ISolverLocalExec_GetExecPath, errorString
End Function

Private Function ISolverLocalExec_CreateSolveScript(ModelFilePathName As String, s As COpenSolver) As String
    ISolverLocalExec_CreateSolveScript = CreateSolveScript(ModelFilePathName, s)
End Function

' ISolverLinear

Private Property Get ISolverLinear_SensitivityAnalysisAvailable() As Boolean
    ISolverLinear_SensitivityAnalysisAvailable = pSensitivityAnalysisAvailable
End Property

' ISolverFile

Private Sub ISolverFile_ReadResults(s As COpenSolver, Optional solution As String)
    ReadResults s
End Sub

Private Property Get ISolverFile_FileType() As OpenSolver_FileType
    ISolverFile_FileType = pFileType
End Property

' Methods

Function SolutionFilePath() As String
    GetTempFilePath pSolutionFile, SolutionFilePath
End Function

Private Function PythonScriptPath() As String
    PythonScriptPath = JoinPaths(ThisWorkbook.Path, SolverDir, pPythonScript)
End Function

Private Function ScriptFilePath() As String
    GetTempFilePath pScriptFile, ScriptFilePath
End Function

Function SensitivityFilePath() As String
    GetTempFilePath pSensitivityFile, SensitivityFilePath
End Function

Function ExecFilePath() As String
    ExecFilePath = JoinPaths(GetBinDir(), pExecName)
End Function

Private Function ShellScriptPath() As String
    #If Mac Then
        ' On Mac, using the gurobi interactive shell causes errors when there are spaces in the filepath.
        ' The mac gurobi.sh script, unlike windows, doesn't have a check for a gurobi install, thus it doesn't do anything for us here and is safe to skip.
        ' We can just run python by itself. We need to use the default system python (pre-installed on mac) and not any other version (e.g. a version from homebrew)
        ' We also need to launch it without going via /Volumes/.../
        ShellScriptPath = JoinPaths(GetRootDriveName(), "usr", "bin", "python")
    #Else
        ShellScriptPath = JoinPaths(GetBinDir(), pShellScript)
    #End If
End Function

Private Function GetGurobiHomeDir() As String
    GetGurobiHomeDir = Environ("GUROBI_HOME")
End Function

Private Function GetBinDir() As String
    #If Mac Then
        GetBinDir = JoinPaths(GetRootDriveName(), "usr", "local", "bin")
    #Else
        GetBinDir = JoinPaths(GetGurobiHomeDir(), "bin")
    #End If
End Function

Function IsPresent(Optional SolverPath As String, Optional errorString As String) As Boolean
    ' Returns true if Gurobi is available and sets SolverPath as path to gurobi_cl
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler
    
    If Not FileOrDirExists(PythonScriptPath()) Then
        errorString = "Unable to find OpenSolver Gurobi script ('" & pPythonScript & "'). Folders searched:" & _
                      vbNewLine & MakePathSafe(JoinPaths(ThisWorkbook.Path, SolverDir))
    #If Win32 Then
        ElseIf Not FileOrDirExists(GetGurobiHomeDir()) Then
            errorString = "No Gurobi installation was detected. " & _
                          "The value of GUROBI_HOME was " & IIf(GetGurobiHomeDir() = "", "not set.", "invalid (" & MakePathSafe(GetGurobiHomeDir()) & ").")
    #End If
    ElseIf Not FileOrDirExists(ExecFilePath()) Then
        errorString = "No Gurobi installation was detected. The following file was not found: " & MakePathSafe(ExecFilePath())
    ElseIf Not FileOrDirExists(ShellScriptPath()) Then
        errorString = "No Gurobi installation was detected. The following file was not found: " & MakePathSafe(ShellScriptPath())
    End If

    If errorString <> "" Then
        IsPresent = False
    Else
        SolverPath = ShellScriptPath()
        IsPresent = True
    End If

ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("CSolverGurobi", "IsPresent") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function

Private Function Version() As String
    ' Get Gurobi version by running 'gurobi_cl -v' at command line
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    If Not SolverIsPresent(Me) Then
        Version = ""
        GoTo ExitFunction
    End If
    
    Dim SolverPath As String
    SolverPath = ExecFilePath()
    
    Dim result As String
    result = ReadExternalCommandOutput(MakePathSafe(SolverPath) & " -v")
    Version = Mid(result, 26, 5)

ExitFunction:
    Close #1
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("CSolverGurobi", "Version") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function

Private Function Bitness() As String
    ' Get Gurobi bitness by running 'gurobi_cl -v' at command line
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    If Not SolverIsPresent(Me) Then
        Bitness = ""
        GoTo ExitFunction
    End If
    
    Dim SolverPath As String
    SolverPath = ExecFilePath()

    Dim result As String
    result = ReadExternalCommandOutput(MakePathSafe(SolverPath) & " -v")
    Bitness = IIf(InStr(result, "64)") > 0, "64", "32")

ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("CSolverGurobi", "Bitness") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function

Private Sub CleanFiles()
    DeleteFileAndVerify SolutionFilePath()
    DeleteFileAndVerify ScriptFilePath()
    DeleteFileAndVerify SensitivityFilePath()
End Sub

Private Function CreateSolveScript(ModelFilePathName As String, s As COpenSolver) As String
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    Dim SolverString As String, CommandLineRunString As String, SolverParametersString As String
    SolverString = MakePathSafe(ShellScriptPath())

    CommandLineRunString = MakePathSafe(PythonScriptPath())
    
    SolverParametersString = ParametersToKwargs(s.SolverParameters)
    
    Dim scriptFile As String, scriptFileContents As String
    scriptFile = ScriptFilePath()
    scriptFileContents = SolverString & " " & CommandLineRunString & " " & SolverParametersString
    CreateScriptFile scriptFile, scriptFileContents
    
    CreateSolveScript = scriptFile

ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("CSolverGurobi", "CreateSolveScript") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function

Sub ReadResults(s As COpenSolver)
    
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    If Not FileOrDirExists(SolutionFilePath()) Then
        Err.Raise Number:=OpenSolver_SolveError, Description:="The solver did not create a solution file. No new solution is available."
    End If
    
    s.SolutionWasLoaded = True
    
    Open SolutionFilePath() For Input As #1 ' supply path with filename
    Dim Line As String, Index As Long
    Line Input #1, Line
    ' Check for python exception while running Gurobi
    Dim GurobiError As String ' The string that identifies a gurobi error in the model file
    GurobiError = "Gurobi Error: "
    If left(Line, Len(GurobiError)) = GurobiError Then
        Err.Raise OpenSolver_GurobiError, Description:=Line
    End If
    'Get the returned status code from gurobi.
    'List of return codes can be seen at - http://www.gurobi.com/documentation/5.1/reference-manual/node865#sec:StatusCodes
    If Line = GurobiResult.Optimal Then
        s.SolveStatus = OpenSolverResult.Optimal
        s.SolveStatusString = "Optimal"
    ElseIf Line = GurobiResult.Infeasible Then
        s.SolveStatus = OpenSolverResult.Infeasible
        s.SolveStatusString = "No Feasible Solution"
        s.SolutionWasLoaded = False
    ElseIf Line = GurobiResult.InfOrUnbound Then
        s.SolveStatus = OpenSolverResult.Unbounded
        s.SolveStatusString = "No Solution Found (Infeasible or Unbounded)"
        s.SolutionWasLoaded = False
    ElseIf Line = GurobiResult.Unbounded Then
        s.SolveStatus = OpenSolverResult.Unbounded
        s.SolveStatusString = "No Solution Found (Unbounded)"
        s.SolutionWasLoaded = False
    ElseIf Line = GurobiResult.SolveStoppedTime Then
        s.SolveStatus = OpenSolverResult.LimitedSubOptimal
        s.SolveStatusString = "Stopped on Time Limit"
    ElseIf Line = GurobiResult.SolveStoppedIter Then
        s.SolveStatus = OpenSolverResult.LimitedSubOptimal
        s.SolveStatusString = "Stopped on Iteration Limit"
    ElseIf Line = GurobiResult.SolveStoppedUser Then
        s.SolveStatus = OpenSolverResult.LimitedSubOptimal
        s.SolveStatusString = "Stopped on Ctrl-C"
    ElseIf Line = GurobiResult.Unsolved Then
        s.SolveStatus = OpenSolverResult.LimitedSubOptimal
        s.SolveStatusString = "Stopped on Gurobi Numerical difficulties"
    ElseIf Line = GurobiResult.SubOptimal Then
        s.SolveStatus = OpenSolverResult.LimitedSubOptimal
        s.SolveStatusString = "Unable to satisfy optimality tolerances; a sub-optimal solution is available."
    Else
        Err.Raise OpenSolver_GurobiError = "The response from the Gurobi solver is not recognised. The response was: " & Line
    End If
    
    If s.SolutionWasLoaded Then
        UpdateStatusBar "OpenSolver: Loading Solution... " & s.SolveStatusString, True
        Dim NumVar As Long, SplitLine() As String
        Line Input #1, Line  ' Optimal - objective value              22
        If Line <> "" Then
            Index = InStr(Line, "=")
            Dim ObjectiveValue As Double
            ObjectiveValue = Val(Mid(Line, Index + 2))
            Dim i As Long
            i = 1
            While Not EOF(1)
                Line Input #1, Line
                SplitLine = SplitWithoutRepeats(Line, " ")
                Dim FinalValue As String
                FinalValue = SplitLine(1)
                
                ' Check for an exponent that is too large, rougly >300
                ' Only do it for -ve exponents, since e-30 ~= e-300
                Index = InStrText(FinalValue, "e-")
                If Index > 0 Then
                    ' Trim the final digit if the exponent has 3 digits
                    If Len(FinalValue) - Index - 1 > 2 Then
                        FinalValue = left(FinalValue, Len(FinalValue) - 1)
                    End If
                End If
                
                s.FinalVarValue(i) = Val(FinalValue)
                s.VarCell(i) = SplitLine(0)
                If left(s.VarCell(i), 1) = "_" Then
                    ' Strip any _ character added to make a valid name
                    s.VarCell(i) = Mid(s.VarCell(i), 2)
                End If
                ' Save number of vars read
                NumVar = i
                i = i + 1
            Wend
        End If
        
        If s.bGetDuals Then
            Open SensitivityFilePath() For Input As 2
            For i = 1 To NumVar
                Line Input #2, Line
                SplitLine = SplitWithoutRepeats(Line, ",")
                s.ReducedCosts(i) = Val(SplitLine(0))
                s.DecreaseVar(i) = s.CostCoeffs(i) - Val(SplitLine(1))
                s.IncreaseVar(i) = Val(SplitLine(2)) - s.CostCoeffs(i)
            Next i

            For i = 1 To s.NumRows
                Line Input #2, Line
                SplitLine = SplitWithoutRepeats(Line, ",")
                s.ShadowPrice(i) = Val(SplitLine(0))
                Dim RHSValue As Double
                RHSValue = Val(SplitLine(1))
                s.IncreaseCon(i) = Val(SplitLine(4)) - RHSValue
                s.DecreaseCon(i) = RHSValue - Val(SplitLine(3))
                s.FinalValue(i) = RHSValue - Val(SplitLine(2))
            Next i
        End If
    End If

ExitSub:
    Application.StatusBar = False
    Close #1
    Close #2
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Sub

ErrorHandler:
    If Not ReportError("CSolverGurobi", "ReadResults") Then Resume
    RaiseError = True
    GoTo ExitSub
End Sub

Private Sub CheckLog(s As COpenSolver)
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    If Not FileOrDirExists(s.LogFilePathName) Then
        Err.Raise Number:=OpenSolver_SolveError, Description:="The solver did not create a log file. No new solution is available."
    End If
    
    Dim message As String
    Open s.LogFilePathName For Input As #1
        message = Input$(LOF(1), 1)
    Close #1

    Dim Key As Variant
    For Each Key In s.SolverParameters.Keys
        If InStrText(message, "No parameters matching '" & Key & "' found") Then
            Err.Raise OpenSolver_SolveError, Description:= _
                "The parameter '" & Key & "' was not recognised by Gurobi. " & _
                "Please check the parameter name you have specified, or consult the Gurobi documentation for more information."
        End If
    Next Key

ExitSub:
    Close #1
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Sub

ErrorHandler:
    If Not ReportError("CSolverGurobi", "CheckLog") Then Resume
    RaiseError = True
    GoTo ExitSub
End Sub
