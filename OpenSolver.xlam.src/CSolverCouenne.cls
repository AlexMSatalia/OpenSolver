VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSolverCouenne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISolver
Implements ISolverLocal
Implements ISolverLocalExec
Implements ISolverFile

Private Const pTitle = "COIN-OR Couenne (Non-linear, non-convex solver)"
Private Const pDesc = "Couenne (Convex Over and Under ENvelopes for Nonlinear Estimation) is a branch & bound algorithm to solve Mixed-Integer Nonlinear Programming (MINLP) problems of specific forms. Couenne aims at finding global optima of nonconvex MINLPs. It implements linearization, bound reduction, and branching methods within a branch-and-bound framework. This solver will fail if your spreadsheet uses functions OpenSolver cannot interpret."
Private Const pLink = "https://projects.coin-or.org/Couenne"
Private Const pName = "Couenne"
Private Const pShortName = "Couenne"

Private Const pExecName = "couenne" & ExecExtension
Private Const pScriptFile = "couenne" & ScriptExtension
Private Const pOptionsFile = "couenne.opt"
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
    DeleteFileAndVerify ScriptFilePath()
    DeleteFileAndVerify OptionsFilePath()
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

Private Function ISolverLocalExec_CreateSolveScript(ModelFilePathName As String, s As COpenSolver) As String
    ISolverLocalExec_CreateSolveScript = CreateSolveScript_NL(ModelFilePathName, s, ScriptFilePath, OptionsFilePath)
End Function

' ISolverFile

Private Sub ISolverFile_ReadResults(s As COpenSolver, Optional solution As String)
    ReadResults_NL s
End Sub

Private Property Get ISolverFile_FileType() As OpenSolver_FileType
    ISolverFile_FileType = pFileType
End Property

' Methods

Function ScriptFilePath() As String
    GetTempFilePath pScriptFile, ScriptFilePath
End Function

Function OptionsFilePath() As String
    GetTempFilePath pOptionsFile, OptionsFilePath
End Function

Private Function Version() As String
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    Dim SolverPath As String
    If Not SolverIsPresent(Me, SolverPath) Then
        Version = ""
        GoTo ExitFunction
    End If
    
    ' Get Couenne version by running 'couenne -v' at command line
    Dim result As String
    result = ReadExternalCommandOutput(MakePathSafe(SolverPath) & " -v")
    Version = Mid(result, 9, 5)

ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("CSolverCouenne", "Version") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function

