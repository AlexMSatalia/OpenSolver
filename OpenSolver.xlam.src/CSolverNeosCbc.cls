VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSolverNeosCbc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISolver
Implements ISolverNeos
Implements ISolverLinear
Implements ISolverFile
Implements ISolverFileAMPL

Private pISolver As ISolver  ' CBC solver to inherit properties from

Private Const pTitle = "NEOS using COIN-OR CBC (Linear solver)"
Private Const pDesc = "The NEOS (Network-Enabled Optimization System) Server is a free Internet-based service for solving optimization problems. The COIN Branch and Cut solver (CBC) is the default solver for OpenSolver and is an open-source mixed-integer program (MIP) solver written in C++. CBC is an active open-source project led by John Forrest at www.coin-or.org."
Private Const pLink = NeosTermsAndConditionsLink
Private Const pShortName = "NeosCBC"

Private Const pSensitivityAnalysisAvailable As Boolean = False

Private Const pNeosSolverName As String = "cbc"
Private Const pNeosSolverCategory As String = "milp"
Private Const pOptionFile As String = vbNullString
Private Const pFileType = OpenSolver_FileType.AMPL
Private Const pModelType = OpenSolver_ModelType.Diff

Private Const pAmplSolverName As String = "cbc"

' Class

Private Sub Class_Initialize()
1         Set pISolver = New CSolverCbc
End Sub

' ISolver

Private Property Get ISolver_Name() As String
1         ISolver_Name = pISolver.Name
End Property

Private Property Get ISolver_ShortName() As String
1         ISolver_ShortName = pShortName
End Property

Private Property Get ISolver_Title() As String
1         ISolver_Title = pTitle
End Property

Private Property Get ISolver_Desc() As String
1         ISolver_Desc = pDesc
End Property

Private Property Get ISolver_Link() As String
1         ISolver_Link = pLink
End Property

Private Property Get ISolver_IterationLimitName() As String
1         ISolver_IterationLimitName = pISolver.IterationLimitName
End Property

Private Property Get ISolver_PrecisionName() As String
1         ISolver_PrecisionName = pISolver.PrecisionName
End Property

Private Property Get ISolver_TimeLimitName() As String
1         ISolver_TimeLimitName = pISolver.TimeLimitName
End Property

Private Property Get ISolver_ToleranceName() As String
1         ISolver_ToleranceName = pISolver.ToleranceName
End Property

Private Property Get ISolver_ModelType() As OpenSolver_ModelType
1         ISolver_ModelType = pModelType
End Property

Private Sub ISolver_CleanFiles()
End Sub

Private Sub ISolver_CheckLog(s As COpenSolver)
1         pISolver.CheckLog s
End Sub

' ISolverNeos

Private Property Get ISolverNeos_NeosSolverName() As String
1         ISolverNeos_NeosSolverName = pNeosSolverName
End Property

Private Property Get ISolverNeos_NeosSolverCategory() As String
1         ISolverNeos_NeosSolverCategory = pNeosSolverCategory
End Property

Private Property Get ISolverNeos_OptionFile() As String
1         ISolverNeos_OptionFile = pOptionFile
End Property

' ISolverLinear

Private Property Get ISolverLinear_SensitivityAnalysisAvailable() As Boolean
1         ISolverLinear_SensitivityAnalysisAvailable = pSensitivityAnalysisAvailable
End Property

' ISolverFile

Private Sub ISolverFile_ReadResults(s As COpenSolver, Optional solution As String)
1         ReadResults_AMPL s, solution
End Sub

Private Property Get ISolverFile_FileType() As OpenSolver_FileType
1         ISolverFile_FileType = pFileType
End Property

' ISolverFileAMPL

Private Property Get ISolverFileAMPL_AmplSolverName() As String
1         ISolverFileAMPL_AmplSolverName = pAmplSolverName
End Property
