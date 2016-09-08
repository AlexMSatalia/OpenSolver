VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSolverBonmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISolver
Implements ISolverLocal
Implements ISolverLocalExec
Implements ISolverFile
Implements ISolverFileNL

Private Const pTitle = "COIN-OR Bonmin (Non-linear solver)"
Private Const pDesc = "Bonmin (Basic Open-source Nonlinear Mixed INteger programming) is an experimental open-source C++ code for solving general MINLPs (Mixed Integer NonLinear Programming). Finds globally optimal solutions to convex nonlinear problems in continuous and discrete variables, and may be applied heuristically to nonconvex problems. Bonmin uses the COIN-OR solvers CBC and IPOPT while solving. For more info on these, see www.coin-or.org/projects. This solver will fail if your spreadsheet uses functions OpenSolver cannot interpret."
Private Const pLink = "https://projects.coin-or.org/Bonmin"
Private Const pName = "Bonmin"
Private Const pShortName = "Bonmin"

Private Const pExecName = "bonmin" & ExecExtension
Private Const pScriptFile = "bonmin" & ScriptExtension
Private Const pOptionFile = "bonmin.opt"
Private Const pFileType = OpenSolver_FileType.NL
Private Const pModelType = OpenSolver_ModelType.Parsed

Private Const pIterationLimitName = "iteration_limit"
Private Const pTimeLimitName = "time_limit"
Private Const pToleranceName = "allowable_fraction_gap"

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
    CleanFiles_NL Me
End Sub

Private Sub ISolver_CheckLog(s As COpenSolver)
    CheckLog_NL s
End Sub

' ISolverLocal

Private Property Get ISolverLocal_Version() As String
    ISolverLocal_Version = Version()
End Property

Private Property Get ISolverLocal_Bitness() As String
    SolverIsPresent Me, Bitness:=ISolverLocal_Bitness
End Property

' ISolverLocalExec

Private Property Get ISolverLocalExec_ExecName() As String
    ISolverLocalExec_ExecName = pExecName
End Property

Private Function ISolverLocalExec_GetExecPath(Optional errorString As String, Optional Bitness As String) As String
    ISolverLocalExec_GetExecPath = SolverExecPath(Me, errorString, Bitness)
End Function

Private Function ISolverLocalExec_CreateSolveCommand(s As COpenSolver) As String
    ISolverLocalExec_CreateSolveCommand = CreateSolveCommand_NL(s, ScriptFilePath)
End Function

' ISolverFile

Private Sub ISolverFile_ReadResults(s As COpenSolver, Optional solution As String)
    ReadResults_NL s
End Sub

Private Property Get ISolverFile_FileType() As OpenSolver_FileType
    ISolverFile_FileType = pFileType
End Property

' ISolverFileNL

Private Property Get ISolverFileNL_OptionFile() As String
    ISolverFileNL_OptionFile = pOptionFile
End Property

' Methods

Function ScriptFilePath() As String
    GetTempFilePath pScriptFile, ScriptFilePath
End Function

Private Function Version() As String
      ' Get Bonmin version by running 'bonmin -v' at command line
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    Dim SolverPath As String
    If Not SolverIsPresent(Me, SolverPath) Then
        Version = ""
        GoTo ExitFunction
    End If
    
    Dim result As String
    result = ExecCapture(MakePathSafe(SolverPath) & " -v")
    Version = Mid(result, 8, 5)

ExitFunction:
    If RaiseError Then RethrowError
    Exit Function

ErrorHandler:
    If Not ReportError("CSolverBonmin", "Version") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function
